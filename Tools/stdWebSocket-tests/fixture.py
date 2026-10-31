import asyncio
from websockets.asyncio.server import serve
from websockets.exceptions import ConnectionClosed

async def handler(ws):
    path = ws.request.path
    if path == '/silent':
        await asyncio.sleep(10)
    elif path == '/idle':
        await asyncio.sleep(0.15)
        await ws.send('idle\0\u20ac')
        await ws.send(b'\0\xff')
        await ws.close(3003, 'async done')
    elif path == '/burst':
        for i in range(100):
            await ws.send(bytes([i]))
        await ws.close(3004, 'burst done')
    elif path == '/async-edge':
        await ws.send(['x' * 6000, 'x' * 4000 + '\0\u20ac'])
        await ws.send('')
        await ws.send(b'')
        await ws.wait_closed()
    elif path in ('/no-close', '/no-read'):
        ws.transport.pause_reading()
        await asyncio.sleep(10)
    elif path == '/fragment':
        await ws.send(['x' * 6000, 'x' * 4000 + '\0\u20ac'])
        await ws.close(3002, 'peer done')
    elif path == '/binary-fragment':
        await ws.send([bytes(range(256)) * 20, b'\0\xff'])
    elif path == '/trickle':
        async with ws.send_context():
            ws.protocol.send_text(b'x', fin=False)
        await asyncio.sleep(10)
    elif path == '/broken-fragment':
        async with ws.send_context():
            ws.protocol.send_text(b'x', fin=False)
        ws.transport.abort()
    elif path == '/split-utf8':
        async with ws.send_context():
            ws.protocol.send_text(b'\xe2', fin=False)
            ws.protocol.send_continuation(b'\x82\xac', fin=True)
    elif path == '/limit':
        await ws.send(b'x' * 16777217)
    elif path == '/tls':
        async for message in ws:
            await ws.send(message)
    elif path == '/echo?test=1':
        try:
            async for message in ws:
                await ws.send(message)
        except ConnectionClosed:
            pass
    else:
        await ws.close(1008, 'unexpected path/query')

async def handshake(connection, request):
    if request.path == '/reject':
        return connection.respond(403, 'Rejected fixture upgrade')
    if request.path == '/stall-handshake':
        await asyncio.sleep(10)

async def start(ready, tls=None):
    async with serve(handler, '127.0.0.1', 0, ssl=tls, process_request=handshake, max_size=32 * 1024 * 1024) as server:
        ready(server.sockets[0].getsockname()[1])
        await asyncio.Future()
