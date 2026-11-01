"""Check both native ABIs without Office. Dev dependencies: keystone-engine, unicorn."""
import struct
import unittest
from build import assemble
from unicorn import Uc, UC_ARCH_X86, UC_MODE_32, UC_MODE_64, UC_HOOK_CODE
from unicorn.x86_const import *


class ThunkTests(unittest.TestCase):
    def check_abi(self, bits, fail=False, mode=1, message=0x8001, remove_success=True, unsafe=False):
        size = bits // 8
        fmt = '<Q' if bits == 64 else '<I'
        pack = lambda value: struct.pack(fmt, value & ((1 << bits) - 1))
        vm = Uc(UC_ARCH_X86, UC_MODE_64 if bits == 64 else UC_MODE_32)
        vm.mem_map(0x100000, 0x10000)
        vm.mem_map(0x200000, 0x10000)
        code = assemble(bits)
        markers = ((0x1122334455667788, 0x2233445566778899) if bits == 64 else (0x11223344, 0x22334455))
        code = code.replace(pack(markers[0]), pack(0x101000)).replace(pack(markers[1]), pack(0x103000))
        mode_marker = 0x33445566778899AA if bits == 64 else 0x33445566
        remove_marker = 0x445566778899AABB if bits == 64 else 0x44556677
        code = code.replace(pack(mode_marker), pack(0x10A000)).replace(pack(remove_marker), pack(0x10B000))
        vm.mem_write(0x100000, code)
        vm.mem_write(0x104000, pack(0x105000))
        vm.mem_write(0x105000 + 7*size, pack(0x106000))
        vm.mem_write(0x105000 + size, pack(0x108000))
        vm.mem_write(0x105000 + 2*size, pack(0x109000))
        sp = 0x208008 if bits == 64 else 0x208000
        reg_sp = UC_X86_REG_RSP if bits == 64 else UC_X86_REG_ESP
        reg_ret = UC_X86_REG_RAX if bits == 64 else UC_X86_REG_EAX
        vm.reg_write(reg_sp, sp)
        # A non-running IDE must not even AddRef this deliberately invalid pointer.
        args = [0x123456, message, (1 << 40) + 123 if bits == 64 else 123, (1 << bits) - 456, 99, (0x104001 if unsafe else (0x104000 if mode == 1 else 0xDEAD0000))]
        vm.mem_write(0x101004, struct.pack('<I', 1))
        vm.mem_write(sp, pack(0x107000))
        if bits == 64:
            for reg, value in zip((UC_X86_REG_RCX, UC_X86_REG_RDX, UC_X86_REG_R8, UC_X86_REG_R9), args):
                vm.reg_write(reg, value)
            vm.mem_write(sp+40, pack(args[4])+pack(args[5]))
        else:
            vm.mem_write(sp+4, b''.join(pack(v) for v in args))
        read = lambda address: struct.unpack(fmt, vm.mem_read(address, size))[0]
        preserved = ((UC_X86_REG_RBX, UC_X86_REG_RBP, UC_X86_REG_RSI, UC_X86_REG_RDI, UC_X86_REG_R12, UC_X86_REG_R13, UC_X86_REG_R14, UC_X86_REG_R15)
                     if bits == 64 else (UC_X86_REG_EBX, UC_X86_REG_EBP, UC_X86_REG_ESI, UC_X86_REG_EDI))
        for reg in preserved:
            vm.reg_write(reg, 0x123000 + reg)
        seen = []

        def hook(vm, address, instruction_size, unused):
            if address not in (0x106000, 0x103000, 0x108000, 0x109000, 0x10A000, 0x10B000):
                return
            stack = vm.reg_read(reg_sp)
            self.assertEqual(int.from_bytes(vm.mem_read(0x101000, 4), 'little'), 1)
            if bits == 64:
                self.assertEqual(stack % 16, 8, 'callee entry alignment')
                incoming = [vm.reg_read(reg) for reg in (UC_X86_REG_RCX, UC_X86_REG_RDX, UC_X86_REG_R8, UC_X86_REG_R9)]
            else:
                incoming = [read(stack+i*4) for i in range(1, 5)]
            if address == 0x10A000:
                seen.append('mode')
                vm.reg_write(reg_ret, mode)
                popped = 0
            elif address == 0x10B000:
                self.assertEqual(incoming[:3], [args[0], 0x100000, args[4]])
                seen.append('detach')
                vm.reg_write(reg_ret, int(remove_success))
                popped = 12 if bits == 32 else 0
            elif address in (0x108000, 0x109000):
                self.assertEqual(incoming[0], 0x104000)
                seen.append('retain' if address == 0x108000 else 'release')
                vm.reg_write(reg_ret, 2 if address == 0x108000 else 1)
                popped = 4 if bits == 32 else 0
            elif address == 0x106000:
                self.assertEqual(incoming[0], 0x104000)
                array = read(incoming[1])
                self.assertEqual(struct.unpack('<HHII', vm.mem_read(array, 12)), (1, 0x812, 24 if bits == 64 else 16, 1))
                data = read(array + (16 if bits == 64 else 12))
                bound = array + (24 if bits == 64 else 16)
                self.assertEqual(struct.unpack('<II', vm.mem_read(bound, 8)), (4, 0))
                for i, expected in enumerate(args[:4]):
                    variant = data+i*(24 if bits == 64 else 16)
                    self.assertEqual(int.from_bytes(vm.mem_read(variant, 2), 'little'), 3 if bits == 32 or i == 1 else 20)
                    actual = int.from_bytes(vm.mem_read(variant+8, 4 if i == 1 else size), 'little')
                    self.assertEqual(actual, expected)
                result = incoming[2]
                vm.mem_write(result, pack(20 if bits == 64 else 3))
                vm.mem_write(result+8, pack(args[2]))
                vm.reg_write(reg_ret, 0x80004005 if fail else 0)
                seen.append('run')
                popped = 12 if bits == 32 else 0
            else:
                self.assertTrue(fail or mode != 1)
                self.assertEqual(incoming, args[:4])
                vm.reg_write(reg_ret, 9876)
                seen.append('forward')
                popped = 16 if bits == 32 else 0
            vm.reg_write(reg_sp, stack+size+popped)
            vm.reg_write(UC_X86_REG_RIP if bits == 64 else UC_X86_REG_EIP, read(stack))

        vm.hook_add(UC_HOOK_CODE, hook)
        vm.emu_start(0x100000, 0x107000)
        if mode == 1 or unsafe:
            expected = ([] if unsafe else ['mode']) + ['retain', 'run', 'release'] + (['forward'] if fail else [])
        else:
            expected = ['mode'] + (['detach'] if mode != 2 or message == 0x82 else []) + ['forward']
        self.assertEqual(seen, expected)
        self.assertEqual(vm.reg_read(reg_ret), 9876 if fail or (mode != 1 and not unsafe) else args[2])
        self.assertEqual(int.from_bytes(vm.mem_read(0x101000, 4), 'little'), 0)
        detached = not unsafe and mode != 1 and (mode != 2 or message == 0x82) and remove_success
        self.assertEqual(int.from_bytes(vm.mem_read(0x101004, 4), 'little'), 0 if detached else 1)
        self.assertEqual(vm.reg_read(reg_sp), sp+size+(24 if bits == 32 else 0))
        for reg in preserved:
            self.assertEqual(vm.reg_read(reg), 0x123000 + reg)

    def test_x86_result(self): self.check_abi(32, False)
    def test_x86_fallback(self): self.check_abi(32, True)
    def test_x64_result(self): self.check_abi(64, False)
    def test_x64_fallback(self): self.check_abi(64, True)
    def test_x86_break(self): self.check_abi(32, mode=2)
    def test_x64_break(self): self.check_abi(64, mode=2)
    def test_x86_design(self): self.check_abi(32, mode=0)
    def test_x64_design(self): self.check_abi(64, mode=0)
    def test_x86_unknown_mode(self): self.check_abi(32, mode=99)
    def test_x64_unknown_mode(self): self.check_abi(64, mode=99)
    def test_x86_break_destroy(self): self.check_abi(32, mode=2, message=0x82)
    def test_x64_break_destroy(self): self.check_abi(64, mode=2, message=0x82)
    def test_x86_remove_failure(self): self.check_abi(32, mode=0, remove_success=False)
    def test_x64_remove_failure(self): self.check_abi(64, mode=0, remove_success=False)
    def test_x86_opt_out_design(self): self.check_abi(32, mode=0, unsafe=True)
    def test_x64_opt_out_design(self): self.check_abi(64, mode=0, unsafe=True)
    def test_x86_opt_out_break(self): self.check_abi(32, mode=2, unsafe=True)
    def test_x64_opt_out_break(self): self.check_abi(64, mode=2, unsafe=True)


if __name__ == '__main__':
    unittest.main()
