Dim C as Long
C = cells(1,1).interior.Color

R = C Mod 256
G = C \ 256 Mod 256
B = C \ 65536 Mod 256
