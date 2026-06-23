"""
Debate Tournament Pool Coder
============================

Deterministic (non-AI-guesswork) logic for two steps:

1. CODING: given a school + category + team count, assign letter codes
   (A, B, C, ...) per school per category. No duplication possible because
   it's a plain loop, not a model "remembering" what it already wrote.

2. POOLING: distribute the coded teams across N pools so that:
   - pool sizes stay as balanced as possible (target ~equal teams/pool)
   - teams from the same school are spread across different pools
     wherever possible (avoid clumping a school into one pool)

Input format (list of dicts):
    {"school": "Jamhuri High School", "category": "Grade 10", "teams": 5}

Output:
    coded_teams: ["Jamhuri High School A", "Jamhuri High School B", ...]
    pools: {"Pool A": [...], "Pool B": [...], ...}
"""

from collections import defaultdict
from string import ascii_uppercase


def code_teams(entries):
    """
    entries: list of {"school": str, "category": str, "teams": int}
    Returns: list of {"school": str, "category": str, "code": str, "team_name": str}
    """
    coded = []
    for entry in entries:
        school = entry["school"]
        category = entry["category"]
        count = entry["teams"]
        if count < 1:
            continue
        for i in range(count):
            letter = ascii_uppercase[i] if i < 26 else f"Z{i}"  # safety net past 26
            team_name = school if count == 1 else f"{school} Team {letter}"
            coded.append({
                "school": school,
                "category": category,
                "code": letter if count > 1 else "",
                "team_name": team_name,
            })
    return coded


def build_pools(coded_teams, num_pools, pool_names=None):
    """
    coded_teams: output of code_teams(), already filtered to one category
    num_pools: how many pools to split into
    pool_names: optional list of pool labels, defaults to "Pool A", "Pool B", ...

    Greedy balancing:
      - Always place the next team into the pool that currently
        (a) has the fewest teams, with ties broken by
        (b) does NOT already contain a team from the same school,
            if such a pool exists among the tied smallest pools.
    """
    if pool_names is None:
        pool_names = [f"Pool {ascii_uppercase[i]}" for i in range(num_pools)]
    assert len(pool_names) == num_pools

    pools = {name: [] for name in pool_names}
    pool_school_sets = {name: set() for name in pool_names}

    # Process schools in round-robin across their own teams first, so a
    # school's teams get spread out over time rather than dumped consecutively.
    by_school = defaultdict(list)
    for t in coded_teams:
        by_school[t["school"]].append(t)

    # Interleave: take one team from each school in turn (round-robin),
    # so large schools don't dominate the early/greedy placements.
    queues = list(by_school.values())
    ordered = []
    while any(queues):
        for q in queues:
            if q:
                ordered.append(q.pop(0))
        queues = [q for q in queues if q]

    for team in ordered:
        school = team["school"]
        min_size = min(len(pools[name]) for name in pool_names)
        smallest_pools = [name for name in pool_names if len(pools[name]) == min_size]

        # Prefer a smallest pool that doesn't already have this school
        candidates = [name for name in smallest_pools if school not in pool_school_sets[name]]
        chosen = candidates[0] if candidates else smallest_pools[0]

        pools[chosen].append(team["team_name"])
        pool_school_sets[chosen].add(school)

    return pools


def run(entries, num_pools, category_filter=None):
    """
    Convenience wrapper: codes all teams, then builds pools per category.
    Returns {category: {"coded": [...], "pools": {...}}}
    """
    coded = code_teams(entries)
    categories = sorted(set(t["category"] for t in coded))
    if category_filter:
        categories = [c for c in categories if c == category_filter]

    result = {}
    for cat in categories:
        cat_teams = [t for t in coded if t["category"] == cat]
        pools = build_pools(cat_teams, num_pools)
        result[cat] = {
            "coded": [t["team_name"] for t in cat_teams],
            "pools": pools,
        }
    return result


if __name__ == "__main__":
    import json
    sample_entries = [
        {"school": "St Anne's Jogoo", "category": "Grade 10", "teams": 4},
        {"school": "Mary Leakey Girls", "category": "Grade 10", "teams": 2},
        {"school": "Jamhuri High School", "category": "Grade 10", "teams": 5},
        {"school": "Moi Forces Academy", "category": "Grade 10", "teams": 2},
        {"school": "Starehe Boys Centre", "category": "Grade 10", "teams": 1},
        {"school": "Makini School", "category": "Grade 10", "teams": 4},
        {"school": "Kirangari Boys", "category": "Grade 10", "teams": 1},
        {"school": "Moi Girls Isinya", "category": "Grade 10", "teams": 2},
        {"school": "St Hannah's", "category": "Grade 10", "teams": 2},
        {"school": "Apostolic Carmel", "category": "Grade 10", "teams": 1},
        {"school": "Ofafa Jericho", "category": "Grade 10", "teams": 1},
        {"school": "Noonkopir Girls", "category": "Grade 10", "teams": 1},
        {"school": "Our Lady of Mercy", "category": "Grade 10", "teams": 1},
    ]
    out = run(sample_entries, num_pools=4)
    print(json.dumps(out, indent=2))
