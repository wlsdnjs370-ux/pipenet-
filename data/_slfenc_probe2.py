import re, io, os

out = io.open('data/_slfenc_probe2.txt', 'w', encoding='utf-8')
p = r'.\data\uploads\module_f\34a12e1d0d3543c6_design\1. 입력도면 대명동 단위세대 평면도_수리계산입력.sdf'
b = open(p, 'rb').read()
out.write('size=%d\n' % len(b))
out.write('head=%r\n' % b[:400])
m = re.search(rb'<User-lib\s+file="([^"]*)"', b)
v = m.group(1).decode('utf-8')
out.write('userlib raw=%r\n' % m.group(1))
out.write('as-utf8=%s\n' % v)
try:
    out.write('latin1->cp949=%s\n' % v.encode('latin-1').decode('cp949'))
except Exception as e:
    out.write('roundtrip fail %s\n' % e)
out.write('\ndir listing:\n')
for f in os.listdir(os.path.dirname(p)):
    out.write(repr(f) + '\n')
out.close()
