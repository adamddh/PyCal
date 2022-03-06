"""script to just run all events, to be used infrequently"""

from pythoncalendar_v3 import run_calvin, check_connection


def main():
    """main()"""
    run_calvin()


if __name__ == '__main__' and check_connection():
    main()
