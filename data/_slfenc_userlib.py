import re, os, io

out = io.open('data/_slfenc_userlib.txt', 'w', encoding='utf-8')
pat = re.compile(rb'<User-lib\s+file="([^"]*)"')
absn = 0
name = 0
name_true = []
name_moj = []
abs_moj = 0
abs_true = 0
SEP = chr(92)
for root, d, fs in os.walk('.'):
    if '.git' in root:
        continue
    for f in fs:
        if not f.lower().endswith('.sdf'):
            continue
        p = os.path.join(root, f)
        b = open(p, 'rb').read()
        for m in pat.finditer(b):
            v = m.group(1).decode('utf-8', errors='replace')
            isabs = (SEP in v) or ('/' in v) or (':' in v)
            na = [c for c in v if ord(c) > 127]
            hi = [c for c in na if ord(c) > 0xFF]
            if isabs:
                absn += 1
                if na:
                    if hi:
                        abs_true += 1
                    else:
                        abs_moj += 1
            else:
                name += 1
                if na:
                    if hi:
                        name_true.append(p)
                    else:
                        name_moj.append(p)
out.write('User-lib absolute: %d (nonascii mojibake %d / true-unicode %d)\n' % (absn, abs_moj, abs_true))
out.write('User-lib filename-only: %d (nonascii true %d / nonascii mojibake %d)\n' % (name, len(name_true), len(name_moj)))
out.write('\nfilename-only + TRUE-unicode examples:\n')
for p in name_true[:6]:
    out.write(p + '\n')
out.write('\nfilename-only + MOJIBAKE examples:\n')
for p in name_moj[:10]:
    out.write(p + '\n')
out.close()
