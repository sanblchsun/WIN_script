import asyncio
from datetime import time
from random import randint


async def execute(delay, value):
    await asyncio.sleep(delay)
    print(value)

async def main():
    # Using asyncio.create_task() method to run coroutines concurrently as asyncio
    for i in range(10):
        slp = randint(1,10)
        task1 = asyncio.create_task(execute(slp, f"{i}  {slp}"))
        await task1




asyncio.run(main())