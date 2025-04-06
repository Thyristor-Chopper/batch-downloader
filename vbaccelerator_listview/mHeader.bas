Attribute VB_Name = "mHeader"
Option Explicit

' Header stuff:
Public Const WC_HEADERA = "SysHeader32"
Public Const WC_HEADER = WC_HEADERA

Public Const HDS_HOTTRACK = &H4 ' v 4.70
Public Const HDS_DRAGDROP = &H40 ' v 4.70
Public Const HDS_FULLDRAG = &H80

Public Const HDI_WIDTH = &H1
Public Const HDI_HEIGHT = HDI_WIDTH
Public Const HDI_TEXT = &H2
Public Const HDI_FORMAT = &H4
Public Const HDI_LPARAM = &H8
Public Const HDI_BITMAP = &H10

'
Public Const HDI_IMAGE = &H20
Public Const HDI_DI_SETITEM = &H40
Public Const HDI_ORDER = &H80

Public Const HDF_LEFT = 0
Public Const HDF_RIGHT = 1
Public Const HDF_CENTER = 2
Public Const HDF_JUSTIFYMASK = &H3
Public Const HDF_RTLREADING = 4
' 4.70+
Public Const HDF_BITMAP_ON_RIGHT = &H1000
Public Const HDF_IMAGE = &H800

Public Const HDF_OWNERDRAW = &H8000
Public Const HDF_STRING = &H4000
Public Const HDF_BITMAP = &H2000

Public Const HDM_FIRST = &H1200                    '// Header messages

Public Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
' Header_GetItemCount(hwndHD) \
'    (int)SendMessage((hwndHD), HDM_GETITEMCOUNT, 0, 0L)
Public Const HDM_INSERTITEMA = (HDM_FIRST + 1)
Public Const HDM_INSERTITEM = HDM_INSERTITEMA
'Header_InsertItem(hwndHD, i, phdi) \
'    (int)SendMessage((hwndHD), HDM_INSERTITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM FAR*)(phdi))
Public Const HDM_DELETEITEM = (HDM_FIRST + 2)
'Header_DeleteItem(hwndHD, i) \
'    (BOOL)SendMessage((hwndHD), HDM_DELETEITEM, (WPARAM)(int)(i), 0L)
Public Const HDM_GETITEMA = (HDM_FIRST + 3)
Public Const HDM_GETITEM = HDM_GETITEMA
'Header_GetItem(hwndHD, i, phdi) \
'    (BOOL)SendMessage((hwndHD), HDM_GETITEM, (WPARAM)(int)(i), (LPARAM)(HD_ITEM FAR*)(phdi))
Public Const HDM_SETITEMA = (HDM_FIRST + 4)
Public Const HDM_SETITEM = HDM_SETITEMA
' Header_SetItem(hwndHD, i, phdi) \
'    (BOOL)SendMessage((hwndHD), HDM_SETITEM, (WPARAM)(int)(i), (LPARAM)(const HD_ITEM FAR*)(phdi))
Public Const HDM_LAYOUT = (HDM_FIRST + 5)
' Header_Layout(hwndHD, playout) \
'    (BOOL)SendMessage((hwndHD), HDM_LAYOUT, 0, (LPARAM)(HD_LAYOUT FAR*)(playout))
Public Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
'  Header_SetImageList(hwnd, himl) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_SETIMAGELIST, 0, (LPARAM)himl)
Public Const HDM_GETIMAGELIST = (HDM_FIRST + 9)
' Header_GetImageList(hwnd) \
'        (HIMAGELIST)SNDMSG((hwnd), HDM_GETIMAGELIST, 0, 0)

Public Const HHT_NOWHERE = &H1
Public Const HHT_ONHEADER = &H2
Public Const HHT_ONDIVIDER = &H4
Public Const HHT_ONDIVOPEN = &H8
Public Const HHT_ABOVE = &H100
Public Const HHT_BELOW = &H200
Public Const HHT_TORIGHT = &H400
Public Const HHT_TOLEFT = &H800
Public Const HDM_HITTEST = (HDM_FIRST + 6)

Public Const H_MAX As Long = &HFFFF + 1
Public Const HDN_FIRST = H_MAX - 300&                  '// header
Public Const HDN_LAST = H_MAX - 399&

Public Const HDN_ITEMCHANGINGA = (HDN_FIRST - 0)
Public Const HDN_ITEMCHANGINGW = (HDN_FIRST - 20)
Public Const HDN_ITEMCHANGEDA = (HDN_FIRST - 1)
Public Const HDN_ITEMCHANGEDW = (HDN_FIRST - 21)
Public Const HDN_ITEMCLICKA = (HDN_FIRST - 2)
Public Const HDN_ITEMCLICKW = (HDN_FIRST - 22)
Public Const HDN_ITEMDBLCLICKA = (HDN_FIRST - 3)
Public Const HDN_ITEMDBLCLICKW = (HDN_FIRST - 23)
Public Const HDN_DIVIDERDBLCLICKA = (HDN_FIRST - 5)
Public Const HDN_DIVIDERDBLCLICKW = (HDN_FIRST - 25)
Public Const HDN_BEGINTRACKA = (HDN_FIRST - 6)
Public Const HDN_BEGINTRACKW = (HDN_FIRST - 26)
Public Const HDN_ENDTRACKA = (HDN_FIRST - 7)
Public Const HDN_ENDTRACKW = (HDN_FIRST - 27)
Public Const HDN_TRACKA = (HDN_FIRST - 8)
Public Const HDN_TRACKW = (HDN_FIRST - 28)
Public Const HDN_ITEMCHANGING = HDN_ITEMCHANGINGA
Public Const HDN_ITEMCHANGED = HDN_ITEMCHANGEDA
Public Const HDN_ITEMCLICK = HDN_ITEMCLICKA
Public Const HDN_ITEMDBLCLICK = HDN_ITEMDBLCLICKA
Public Const HDN_DIVIDERDBLCLICK = HDN_DIVIDERDBLCLICKA
Public Const HDN_BEGINTRACK = HDN_BEGINTRACKA
Public Const HDN_ENDTRACK = HDN_ENDTRACKA
Public Const HDN_TRACK = HDN_TRACKA

' v 4.70
Public Const HDN_BEGINDRAG = (HDN_FIRST - 10)
Public Const HDN_ENDDRAG = (HDN_FIRST - 11)

Public Type HD_HITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
End Type

Public Type HD_ITEM
   mask As Long
   cxy As Long
   pszText As String
   hbm As Long
   cchTextMax As Long
   fmt As Long
   lParam As Long
   ' 4.70:
   iImage As Long
   iOrder As Long
End Type


