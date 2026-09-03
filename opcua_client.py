"""Async event loop runner and OPC UA connection wrapper."""

import asyncio
import threading

from asyncua import Client

from data_io import s


class Async:
    def __init__(self):
        self.loop = asyncio.new_event_loop()
        threading.Thread(target=self.run, daemon=True).start()

    def run(self):
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def call(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop).result(15)

    def stop(self):
        self.loop.call_soon_threadsafe(self.loop.stop)


class Conn:
    def __init__(self, cfg):
        self.cfg = cfg
        self.client = None

    async def connect(self):
        ip = s(self.cfg.get("IP"))
        port = int(self.cfg.get("Port") or 4840)
        ep = s(self.cfg.get("EndpointPath"))
        if ep and not ep.startswith("/"):
            ep = "/" + ep
        self.client = Client(f"opc.tcp://{ip}:{port}{ep}")
        if s(self.cfg.get("Username")):
            self.client.set_user(s(self.cfg["Username"]))
            self.client.set_password(s(self.cfg.get("Password")))
        await self.client.connect()

    async def read(self, node):
        return await self.client.get_node(node).read_value()

    async def write(self, node, value):
        n = self.client.get_node(node)
        dv = await n.read_data_value()
        from asyncua import ua
        await n.write_value(ua.DataValue(ua.Variant(value, dv.Value.VariantType)))

    async def disconnect(self):
        if self.client:
            await self.client.disconnect()
