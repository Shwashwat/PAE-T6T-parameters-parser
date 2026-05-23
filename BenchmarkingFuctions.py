from time import perf_counter

def timer(func):

    def inner(*args, **kwargs):
        start = perf_counter()
        result = func(*args, **kwargs)
        inner.last_elapsed = perf_counter() - start
        print(f'time taken = {inner.last_elapsed}')
        return result

    inner.last_elapsed = 0.0
    return inner