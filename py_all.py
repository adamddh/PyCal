"""script to just run all events, to be used infrequently"""

from pythoncalendar_v3 import ARG, check_connection, run_calvin


def main():
    """main()"""
    run_calvin()


if __name__ == '__main__' and check_connection(ARG):
    main()
