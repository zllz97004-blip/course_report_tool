DEBUG = True
ATTAINMENT_THRESHOLD = 0.60


def debug_print(*args, **kwargs):
    if DEBUG:
        print(*args, **kwargs)
