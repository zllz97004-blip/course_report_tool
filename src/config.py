DEBUG = True


def debug_print(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)
