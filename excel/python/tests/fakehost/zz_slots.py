"""Test only: 2,000 worksheet functions, which with checks.py's overfill the add-in's pool of 2,000
slots. Named so that it imports after checks.py, whose functions then take the first slots.

slot_NNNN takes NNNN % 4 arguments and returns NNNN * 100 plus their sum, so a call shows
both that the right function runs in a slot and that the slot received its arguments."""
from asfpy import xl_func

for _i in range(2000):
    _names = ['a%d' % k for k in range(_i % 4)]
    exec('def slot_%04d(%s):\n    return %d + %s\n'
         % (_i, ', '.join(_names), _i * 100, ' + '.join(['0'] + _names)))
    xl_func(globals()['slot_%04d' % _i])
