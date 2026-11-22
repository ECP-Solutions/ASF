"""Test only, for the fake host: 1,000 worksheet functions, slot_0000 to slot_0999. run_fakehost.py
also writes a copy for slot_1000 to slot_1999 (zz_slots_more.py), so that with checks.py's they
overfill the add-in's pool of 2,000 slots. Named so that it imports after checks.py, whose
functions then take the first slots. Not for the Excel checklist's Python folder.

slot_NNNN takes NNNN % 4 arguments and returns NNNN * 100 plus their sum, so a call shows
both that the right function runs in a slot and that the slot received its arguments."""
from asfpy import xl_func

FIRST, LAST = 0, 1000

for _i in range(FIRST, LAST):
    _names = ['a%d' % k for k in range(_i % 4)]
    exec('def slot_%04d(%s):\n    return %d + %s\n'
         % (_i, ', '.join(_names), _i * 100, ' + '.join(['0'] + _names)))
    xl_func(globals()['slot_%04d' % _i])
