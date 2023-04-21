'ExcelPython
'Tieng Viet Unicode - Kieu go VNI
'Dev: Nguyen Thanh Dong

Function TiengVietVNI(text)
arr_unicode = Array(ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), ChrW(7905), ChrW(7907), ChrW(225), ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), ChrW(233), ChrW(232), ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(250), ChrW(249), ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925), ChrW(273))
arr_go_vni = Array("e61", "e62", "e63", "e64", "e65", "a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "u71", "u72", "u73", "u74", "u75", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", "o75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "i1", "i2", "i3", "i4", "i5", "e1", "e2", "e3", "e4", "e5", "e6", "u1", "u2", "u3", "u4", "u5", "u7", "o1", "o2", "o3", "o4", "o5", "o6", "o7", "y1", "y2", "y3", "y4", "y5", "d9")
For i = 0 To UBound(arr_go_vni)
    text = Replace(text, arr_go_vni(i), arr_unicode(i))
    text = Replace(text, UCase(arr_go_vni(i)), UCase(arr_unicode(i)))
Next i
TiengVietVNI = text
End Function
