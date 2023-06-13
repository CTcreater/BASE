import asyncio
loop= asyncio.get_event_loop()
async def work1():
    for x in range(5):
        print("work1")
        await asyncio.sleep(1)

async def work2():
    for x in range(5):
        print("work2")
        await asyncio.sleep(1)
    return x
#task=[work1(),work2()]

async def main():
    task= [asyncio.create_task(work1()),asyncio.create_task(work2())]
    a=await asyncio.wait(task)
    b=await asyncio.gather(*task)
    
    return b
a=asyncio.run(main())
print(a)
