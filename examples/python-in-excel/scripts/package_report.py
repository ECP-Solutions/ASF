"""Run Python script: which Python runs inside Excel, and which packages it has, with versions."""
import sys
from importlib import metadata

print('Python %s, %d-bit' % (sys.version.split()[0], 64 if sys.maxsize > 2 ** 32 else 32))
print(sys.prefix)
print()
for dist in sorted(metadata.distributions(), key=lambda d: d.metadata['Name'].lower()):
    print('  %-22s %s' % (dist.metadata['Name'], dist.version))
