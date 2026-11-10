"""Check vtable slots against Microsoft headers and both native structure layouts.

python tests/native/test_outlook_abi.py --download
python tests/native/test_outlook_abi.py --headers path/to/MAPIStubLibrary/include
Only --download accesses the network; downloaded blobs are verified by Git SHA.
"""
import argparse
import base64
import ctypes as c
import hashlib
import json
from pathlib import Path
import re
import urllib.request

ROOT = Path(__file__).resolve().parents[2]
BLOBS = {
    "MAPIDefS.h": "73fd58cd08ea27d3595a94594a7e81b26660be80",
    "MAPIX.h": "b7597f57c36e2cfd4273d980b204dbf310cda0df",
}


def read_headers(args):
    result = {}
    for name, sha in BLOBS.items():
        if args.download:
            request = urllib.request.Request(
                f"https://api.github.com/repos/microsoft/MAPIStubLibrary/git/blobs/{sha}",
                headers={"User-Agent": "stdOutlook-ABI-check", "Accept": "application/vnd.github+json"},
            )
            with urllib.request.urlopen(request, timeout=30) as response:
                data = base64.b64decode(json.load(response)["content"])
            digest = hashlib.sha1(b"blob " + str(len(data)).encode() + b"\0" + data).hexdigest()
            if digest != sha:
                raise RuntimeError(f"Git blob verification failed: {name}")
        else:
            data = (args.headers / name).read_bytes()
        result[name] = data.decode("utf-8-sig")
    return result


def methods(header, macro):
    match = re.search(r"#define\s+" + macro + r"\(", header)
    if not match:
        raise AssertionError(f"Header macro missing: {macro}")
    block = header[match.start():].split("#undef", 1)[0]
    names = re.findall(r"MAPIMETHOD(?:_\(\w+,\s*|\()(\w+)\)", block)
    if not names:
        raise AssertionError(f"No methods in {macro}")
    return names


def verify_slots(headers, source):
    defs, ext = headers["MAPIDefS.h"], headers["MAPIX.h"]
    unknown = methods(defs, "MAPI_IUNKNOWN_METHODS")
    prop = unknown + methods(defs, "MAPI_IMAPIPROP_METHODS")
    interfaces = {
        "unknown": unknown,
        "prop": prop,
        "status": prop + methods(defs, "MAPI_IMAPISTATUS_METHODS"),
        "message": prop + methods(defs, "MAPI_IMESSAGE_METHODS"),
        "folder": prop + methods(defs, "MAPI_IMAPICONTAINER_METHODS") + methods(defs, "MAPI_IMAPIFOLDER_METHODS"),
        "session": unknown + methods(ext, "MAPI_IMAPISESSION_METHODS"),
        "book": prop + methods(ext, "MAPI_IADDRBOOK_METHODS"),
        "container": prop + methods(defs, "MAPI_IMAPICONTAINER_METHODS"),
        "store": prop + methods(defs, "MAPI_IMSGSTORE_METHODS"),
        "table": unknown + methods(defs, "MAPI_IMAPITABLE_METHODS"),
    }
    targets = {
        "VT_RELEASE": ("unknown", "Release"),
        "VT_GET_PROPS": ("prop", "GetProps"),
        "VT_SET_PROPS": ("prop", "SetProps"),
        "VT_NAMED_IDS": ("prop", "GetIDsFromNames"),
        "VT_AB_ONE_OFF": ("book", "CreateOneOff"),
        "VT_FOLDER_CREATE_MESSAGE": ("folder", "CreateMessage"),
        "VT_MESSAGE_RECIPIENTS": ("message", "ModifyRecipients"),
        "VT_MESSAGE_SUBMIT": ("message", "SubmitMessage"),
        "VT_SESSION_STORES": ("session", "GetMsgStoresTable"),
        "VT_SESSION_OPEN_STORE": ("session", "OpenMsgStore"),
        "VT_SESSION_ADDRESS_BOOK": ("session", "OpenAddressBook"),
        "VT_SESSION_STATUS_TABLE": ("session", "GetStatusTable"),
        "VT_SESSION_OPEN_ENTRY": ("session", "OpenEntry"),
        "VT_SESSION_COMPARE_IDS": ("session", "CompareEntryIDs"),
        "VT_STATUS_FLUSH": ("status", "FlushQueues"),
        "VT_SESSION_IDENTITY": ("session", "QueryIdentity"),
        "VT_AB_OPEN_ENTRY": ("book", "OpenEntry"),
        "VT_CONTAINER_CONTENTS": ("container", "GetContentsTable"),
        "VT_CONTAINER_HIERARCHY": ("container", "GetHierarchyTable"),
        "VT_STORE_OPEN_ENTRY": ("store", "OpenEntry"),
        "VT_STORE_RECEIVE_FOLDER": ("store", "GetReceiveFolder"),
        "VT_TABLE_COLUMNS": ("table", "SetColumns"),
        "VT_TABLE_RESTRICT": ("table", "Restrict"),
        "VT_TABLE_SORT": ("table", "SortTable"),
        "VT_TABLE_ROWS": ("table", "QueryRows"),
    }
    actual = dict(re.findall(r"Private Const (VT_\w+) As Long = (\d+)", source))
    assert set(actual) == set(targets), "Update ABI coverage when native methods change"
    for constant, (interface, method) in targets.items():
        expected = interfaces[interface].index(method)
        assert int(actual[constant]) == expected, f"{constant}: {actual[constant]} != {expected}"
    print(f"PASS: {len(targets)} vtable slots match Microsoft's MAPI headers")


def verify_layouts(source):
    # Explicit pointer-width fields let either Python architecture model both
    # SDK layouts. This verifies layout, not execution in 32-bit Office.
    block = source.split("#If Win64 Then", 1)[1].split("#End If", 1)[0]
    for ptr, branch in zip((c.c_uint64, c.c_uint32), block.split("#Else")):
        u32 = c.c_uint32

        class Binary(c.Structure):
            _fields_ = [("count", u32), ("pointer", ptr)]

        class Value(c.Union):
            _fields_ = [("binary", Binary), ("number", c.c_double)]

        class Prop(c.Structure):
            _fields_ = [("tag", u32), ("pad", u32), ("value", Value)]

        class Row(c.Structure):
            _fields_ = [("pad", u32), ("count", u32), ("props", ptr)]

        class RowSet(c.Structure):
            _fields_ = [("count", u32), ("rows", Row * 1)]

        class PropertyRestriction(c.Structure):
            _fields_ = [("relop", u32), ("tag", u32), ("prop", ptr)]

        class CommentRestriction(c.Structure):
            _fields_ = [("count", u32), ("restriction", ptr), ("prop", ptr)]

        class RestrictionUnion(c.Union):
            _fields_ = [("property", PropertyRestriction), ("comment", CommentRestriction)]

        class Restriction(c.Structure):
            _fields_ = [("kind", u32), ("value", RestrictionUnion)]

        class NameID(c.Structure):
            _fields_ = [("guid", ptr), ("kind", u32), ("id_or_name", ptr)]

        expected = {
            "PTR_BYTES": c.sizeof(ptr), "PROP_BYTES": c.sizeof(Prop),
            "ROW_BYTES": c.sizeof(Row), "ROWSET_OFFSET": RowSet.rows.offset,
            "RESTRICTION_BYTES": c.sizeof(Restriction),
            "RESTRICTION_UNION": Restriction.value.offset,
            "BINARY_POINTER_OFFSET": Binary.pointer.offset,
        }
        actual = {key: int(value) for key, value in re.findall(r"Private Const (\w+) As Long = (\d+)", branch)}
        assert actual == expected, f"Native layout mismatch: {actual} != {expected}"
        assert Prop.value.offset == 8 and Row.props.offset == 8
        assert PropertyRestriction.prop.offset == 8
        assert c.sizeof(NameID) == c.sizeof(ptr) * 3
        assert NameID.id_or_name.offset == c.sizeof(ptr) * 2
        print(f"PASS: {c.sizeof(ptr) * 8}-bit structure sizes and field offsets")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--download", action="store_true")
    group.add_argument("--headers", type=Path)
    args = parser.parse_args()
    source = (ROOT / "src/stdOutlook.cls").read_text()
    verify_slots(read_headers(args), source)
    verify_layouts(source)


if __name__ == "__main__":
    main()
