import asyncio, time
from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy import text

engine = create_async_engine(
    'mysql+aiomysql://chat_bot:password@74.225.249.46:3307/vmail?charset=utf8mb4',
    pool_size=5, pool_pre_ping=True
)
Session = async_sessionmaker(engine, class_=AsyncSession, expire_on_commit=False)

async def test():
    # First call warms the pool
    t0 = time.time()
    async with Session() as s:
        await s.execute(text('SELECT 1'))
    t1 = time.time()

    # Second call should reuse pooled connection
    async with Session() as s:
        await s.execute(text('SELECT 1'))
    t2 = time.time()

    # Third call
    async with Session() as s:
        await s.execute(text('SELECT 1'))
    t3 = time.time()

    # Simulate typical dashboard: /me + /documents/ sequential
    async with Session() as s:
        await s.execute(text('SELECT id, name, email FROM users WHERE id = 95'))
    t4 = time.time()

    async with Session() as s:
        await s.execute(text('SELECT id, title, doc_type, updated_at, is_trashed, owner_id FROM documents WHERE owner_id = 95'))
    t5 = time.time()

    print(f"1st (cold):  {t1-t0:.3f}s")
    print(f"2nd (warm):  {t2-t1:.3f}s")
    print(f"3rd (warm):  {t3-t2:.3f}s")
    print(f"/me query:   {t4-t3:.3f}s")
    print(f"/docs query: {t5-t4:.3f}s")
    print(f"Total:       {t5-t0:.3f}s")

    await engine.dispose()

asyncio.run(test())
