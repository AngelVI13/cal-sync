from ics import Calendar, Event
from typing import Dict
from pathlib import Path
from ipdb import set_trace

CAL_1 = Path("./calendar1.ics")
CAL_2 = Path("./calendar2.ics")
CAL_3 = Path("./calendar3.ics")


def is_time_range_different(old_ev: Event, new_ev: Event) -> bool:
    old_begin = old_ev.begin.format()
    new_begin = new_ev.begin.format()

    old_end = old_ev.end.format()
    new_end = new_ev.end.format()

    if old_begin != new_begin:
        print(f"event {old_ev.name!r} start changed")
        print(f"{old_begin} -> {new_begin}")

    if old_end != new_end:
        print(f"event {old_ev.name!r} end changed")
        print(f"{old_end} -> {new_end}")
    return False


def compare(old: Dict[int, Event], new: Dict[int, Event]) -> None:
    for hash_, ev in new.items():
        if hash_ not in old:
            print("new event")
            print(ev.name)
            continue

        if is_time_range_different(old[hash_], ev):
            print(f"time changed for event: {ev.name}")


def main():
    c1 = Calendar(CAL_1.read_text())
    c2 = Calendar(CAL_2.read_text())
    c3 = Calendar(CAL_3.read_text())
    events1 = {hash(ev): ev for ev in c1.events}
    events2 = {hash(ev): ev for ev in c2.events}
    events3 = {hash(ev): ev for ev in c3.events}

    compare(events2, events3)


if __name__ == "__main__":
    main()
