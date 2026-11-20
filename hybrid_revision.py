"""Synthetic revision validation; not connected to business records."""
def next_revision(current, expected):
    if type(current) is not int or type(expected) is not int or min(current, expected)<0:
        raise ValueError("Invalid revision")
    if current != expected:
        raise RuntimeError("Revision conflict")
    return current + 1
