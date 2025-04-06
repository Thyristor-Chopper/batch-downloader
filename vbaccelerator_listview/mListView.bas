Attribute VB_Name = "mListView"
Option Explicit


' =========================================================================================================================================
' CC and LV declares
' =========================================================================================================================================
Public Const ICC_LISTVIEW_CLASSES = &H1& ' ' listview, header
Public Const ODT_LISTVIEW = &H102
Public Const LVM_FIRST = &H1000&                   '' ListView messages
Public Const LVN_FIRST = -100                     '' listview
Public Const LVN_LAST = -199

Public Declare Function InitCommonControlsEx Lib "Comctl32.dll" (icc As ICCEx) As Long
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()
Public Type ICCEx
    dwSize As Long          ' size of this structure
    dwICC As Long           ' flags indicating which classes to be initialized
End Type

Public Const CCM_FIRST = &H2000&                   '// Common control shared messages
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)         '// lParam is bkColor

Public Type COLORSCHEME
   dwSize As Long ';
   clrBtnHighlight As Long ';       // highlight color
   clrBtnShadow As Long ';          // shadow color
End Type

Public Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)     '// lParam is color scheme
Public Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)     '// fills in COLORSCHEME pointed to by lParam
Public Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
'#if (_WIN32_IE >= = &H0500)
'public Const COMCTL32_VERSION = 5
Public Const CCM_SETVERSION = (CCM_FIRST + 7)
Public Const CCM_GETVERSION = (CCM_FIRST + 8)
Public Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + 9)    '// wParam == hwndParent.

' Notification messages.
Public Const NM_FIRST = 0
Public Const NM_CLICK = (NM_FIRST - 2)
Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Public Const NM_DBLCLK = (NM_FIRST - 3)
Public Const NM_KILLFOCUS = (NM_FIRST - 8)
Public Const NM_RCLICK = (NM_FIRST - 5)
Public Const NM_RETURN = (NM_FIRST - 4)

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long ' this is control specific, but it's how to specify an item.  valid only with CDDS_ITEM bit set
    uItemState As Long
    lItemlParam As Long
End Type

' CustomDraw paint stages.
Public Const CDDS_PREPAINT = &H1
Public Const CDDS_POSTPAINT = &H2
Public Const CDDS_PREERASE = &H3
Public Const CDDS_POSTERASE = &H4
Public Const CDDS_ITEMPREPAINT = (&H10000 Or &H1)
Public Const CDDS_ITEMPOSTPAINT = (&H10000 Or &H2)
Public Const CDDS_ITEM = &H10000
Public Const CDDS_SUBITEM = &H20000

' CustomDraw Item states. Only the ones we need.
Public Const CDIS_CHECKED = &H8
Public Const CDIS_FOCUS = &H10
Public Const CDIS_HOT = &H40

' CustomDraw return values.
Public Const CDRF_DODEFAULT = &H0
Public Const CDRF_NEWFONT = &H2
Public Const CDRF_SKIPDEFAULT = &H4
Public Const CDRF_NOTIFYPOSTPAINT = &H10
Public Const CDRF_NOTIFYITEMDRAW = &H20
Public Const CDRF_NOTIFYPOSTERASE = &H40
Public Const CDRF_NOTIFYSUBITEMDRAW = &H20

' Header control styles
Public Const HDS_HOTTRACK = &H4 ' v 4.70
Public Const HDS_BUTTONS = &H2


'====== LISTVIEW CONTROL =====================================================

' #ifndef NOLISTVIEW

' #ifdef _WIN32

Public Const WC_LISTVIEWA = "SysListView32"
'public const WC_LISTVIEWW            L"SysListView32"

#If UNICODE Then
Public Const WC_LISTVIEW = WC_LISTVIEWW
#Else
Public Const WC_LISTVIEW = WC_LISTVIEWA
#End If

' begin_r_commctrl

Public Const LVS_ICON = &H0
Public Const LVS_REPORT = &H1
Public Const LVS_SMALLICON = &H2
Public Const LVS_LIST = &H3
Public Const LVS_TYPEMASK = &H3

Public Const LVS_SINGLESEL = &H4
Public Const LVS_SHOWSELALWAYS = &H8
Public Const LVS_SORTASCENDING = &H10
Public Const LVS_SORTDESCENDING = &H20
Public Const LVS_SHAREIMAGELISTS = &H40
Public Const LVS_NOLABELWRAP = &H80
Public Const LVS_AUTOARRANGE = &H100
Public Const LVS_EDITLABELS = &H200
' #if (_WIN32_IE >= =&H0300)
Public Const LVS_OWNERDATA = &H1000
' #end If
Public Const LVS_NOSCROLL = &H2000

Public Const LVS_TYPESTYLEMASK = &HFC00

Public Const LVS_ALIGNTOP = &H0
Public Const LVS_ALIGNLEFT = &H800
Public Const LVS_ALIGNMASK = &HC00

Public Const LVS_OWNERDRAWFIXED = &H400
Public Const LVS_NOCOLUMNHEADER = &H4000
Public Const LVS_NOSORTHEADER = &H8000

' end_r_commctrl

' #if (_WIN32_IE >= =&H0400)
Public Const LVM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
'public const ListView_SetUnicodeFormat(hwnd, fUnicode)  \
'    (BOOL)SNDMSG((hwnd), LVM_SETUNICODEFORMAT, (WPARAM)(fUnicode), 0)

Public Const LVM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
'public const ListView_GetUnicodeFormat(hwnd)  \
'    (BOOL)SNDMSG((hwnd), LVM_GETUNICODEFORMAT, 0, 0)
' #end If

Public Const LVM_GETBKCOLOR = (LVM_FIRST + 0)
'public const ListView_GetBkColor(hwnd)  \
'    (COLORREF)SNDMSG((hwnd), LVM_GETBKCOLOR, 0, 0L)

Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
'public const ListView_SetBkColor(hwnd, clrBk) \
'    (BOOL)SNDMSG((hwnd), LVM_SETBKCOLOR, 0, (LPARAM)(COLORREF)(clrBk))

Public Const LVM_GETIMAGELIST = (LVM_FIRST + 2)
'public const ListView_GetImageList(hwnd, iImageList) \
'    (HIMAGELIST)SNDMSG((hwnd), LVM_GETIMAGELIST, (WPARAM)(INT)(iImageList), 0L)

Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1
Public Const LVSIL_STATE = 2

Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
'public const ListView_SetImageList(hwnd, himl, iImageList) \
'    (HIMAGELIST)SNDMSG((hwnd), LVM_SETIMAGELIST, (WPARAM)(iImageList), (LPARAM)(HIMAGELIST)(himl))

Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
'public const ListView_GetItemCount(hwnd) \
'    (int)SNDMSG((hwnd), LVM_GETITEMCOUNT, 0, 0L)


Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8
' #if (_WIN32_IE >= =&H0300)
Public Const LVIF_INDENT = &H10
'#if (_WIN32_WINNT >= 0x501)
Public Const LVIF_GROUPID = &H100
Public Const LVIF_COLUMNS = &H200
' #end if

Public Const LVIF_NORECOMPUTE = &H800
' #end If

Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_CUT = &H4
Public Const LVIS_DROPHILITED = &H8
Public Const LVIS_ACTIVATING = &H20

Public Const LVIS_OVERLAYMASK = &HF00
Public Const LVIS_STATEIMAGEMASK = &HF000

'public const INDEXTOSTATEIMAGEMASK(i) ((i) << 12)

' #if (_WIN32_IE >= =&H0300)
Public Const I_INDENTCALLBACK = (-1)
'public Const LV_ITEMA = LVITEMA
'public Const LV_ITEMW = LVITEMW
' #else
'public const tagLVITEMA    _LV_ITEMA
'public const LVITEMA       LV_ITEMA
'public const tagLVITEMW    _LV_ITEMW
'public const LVITEMW       LV_ITEMW
' #end If

'public const LV_ITEM LVITEM

'public const LVITEMA_V1_SIZE CCSIZEOF_STRUCT(LVITEMA, lParam)
'public const LVITEMW_V1_SIZE CCSIZEOF_STRUCT(LVITEMW, lParam)

#If UNICODE Then
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
' #if (_WIN32_IE >= =&H0300)
    iIndent As Long
' #end If
End Type
#Else
Public Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
' #if (_WIN32_IE >= =&H0300)
    iIndent As Long
' #end If
'#if (_WIN32_WINNT >= 0x501)
    iGroupId As Long
    cColumns As Long '; // tile view columns
    puColumns As Long
'#End If
End Type
#End If
Public Type LVITEM_LT
   mask As Long
   iItem As Long
   iSubItem As Long
   state As Long
   stateMask As Long
   pszText As Long
   cchTextMax As Long
   iImage As Long
   lParam As Long
   iIndent As Long
End Type

'public const LPSTR_TEXTCALLBACKW     ((LPWSTR)-1L)
Public Const LPSTR_TEXTCALLBACKA = -1&  '  ((LPSTR)-1L)
#If UNICODE Then
Public Const LPSTR_TEXTCALLBACK = LPSTR_TEXTCALLBACKW
#Else
Public Const LPSTR_TEXTCALLBACK = LPSTR_TEXTCALLBACKA
#End If

Public Const I_IMAGECALLBACK = (-1)

Public Const LVM_GETITEMA = (LVM_FIRST + 5)
Public Const LVM_GETITEMW = (LVM_FIRST + 75)
#If UNICODE Then
   Public Const LVM_GETITEM = LVM_GETITEMW
#Else
   Public Const LVM_GETITEM = LVM_GETITEMA
#End If

'public const ListView_GetItem(hwnd, pitem) \
'    (BOOL)SNDMSG((hwnd), LVM_GETITEM, 0, (LPARAM)(LV_ITEM FAR*)(pitem))


Public Const LVM_SETITEMA = (LVM_FIRST + 6)
Public Const LVM_SETITEMW = (LVM_FIRST + 76)
#If UNICODE Then
Public Const LVM_SETITEM = LVM_SETITEMW
#Else
Public Const LVM_SETITEM = LVM_SETITEMA
#End If

'public const ListView_SetItem(hwnd, pitem) \
'    (BOOL)SNDMSG((hwnd), LVM_SETITEM, 0, (LPARAM)(const LV_ITEM FAR*)(pitem))


Public Const LVM_INSERTITEMA = (LVM_FIRST + 7)
Public Const LVM_INSERTITEMW = (LVM_FIRST + 77)
#If UNICODE Then
Public Const LVM_INSERTITEM = LVM_INSERTITEMW
#Else
Public Const LVM_INSERTITEM = LVM_INSERTITEMA
#End If
'public const ListView_InsertItem(hwnd, pitem)   \
'    (int)SNDMSG((hwnd), LVM_INSERTITEM, 0, (LPARAM)(const LV_ITEM FAR*)(pitem))


Public Const LVM_DELETEITEM = (LVM_FIRST + 8)
'public const ListView_DeleteItem(hwnd, i) \
'    (BOOL)SNDMSG((hwnd), LVM_DELETEITEM, (WPARAM)(int)(i), 0L)


Public Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)
'public const ListView_DeleteAllItems(hwnd) \
'    (BOOL)SNDMSG((hwnd), LVM_DELETEALLITEMS, 0, 0L)


Public Const LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
'public const ListView_GetCallbackMask(hwnd) \
'    (BOOL)SNDMSG((hwnd), LVM_GETCALLBACKMASK, 0, 0)


Public Const LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
'public const ListView_SetCallbackMask(hwnd, mask) \
'    (BOOL)SNDMSG((hwnd), LVM_SETCALLBACKMASK, (WPARAM)(UINT)(mask), 0)


Public Const LVNI_ALL = &H0
Public Const LVNI_FOCUSED = &H1
Public Const LVNI_SELECTED = &H2
Public Const LVNI_CUT = &H4
Public Const LVNI_DROPHILITED = &H8

Public Const LVNI_ABOVE = &H100
Public Const LVNI_BELOW = &H200
Public Const LVNI_TOLEFT = &H400
Public Const LVNI_TORIGHT = &H800


Public Const LVM_GETNEXTITEM = (LVM_FIRST + 12)
'public const ListView_GetNextItem(hwnd, i, flags) \
'    (int)SNDMSG((hwnd), LVM_GETNEXTITEM, (WPARAM)(int)(i), MAKELPARAM((flags), 0))


Public Const LVFI_PARAM = &H1
Public Const LVFI_STRING = &H2
Public Const LVFI_PARTIAL = &H8
Public Const LVFI_WRAP = &H20
Public Const LVFI_NEARESTXY = &H40

#If UNICODE Then
Public Type LVFINDINFO
    flags As Long
    psz As Long
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type
#Else
Public Type LVFINDINFO
    flags As Long
    psz As String
    lParam As Long
    pt As POINTAPI
    vkDirection As Long
End Type
#End If


Public Const LVM_FINDITEMA = (LVM_FIRST + 13)
Public Const LVM_FINDITEMW = (LVM_FIRST + 83)
#If UNICODE Then
Public Const LVM_FINDITEM = LVM_FINDITEMW
#Else
Public Const LVM_FINDITEM = LVM_FINDITEMA
#End If

'public const ListView_FindItem(hwnd, iStart, plvfi) \
'    (int)SNDMSG((hwnd), LVM_FINDITEM, (WPARAM)(int)(iStart), (LPARAM)(const LV_FINDINFO FAR*)(plvfi))

Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
Public Const LVIR_SELECTBOUNDS = 3


Public Const LVM_GETITEMRECT = (LVM_FIRST + 14)
'public const ListView_GetItemRect(hwnd, i, prc, code) \
'     (BOOL)SNDMSG((hwnd), LVM_GETITEMRECT, (WPARAM)(int)(i), \
'           ((prc) ? (((RECT FAR *)(prc))->left = (code),(LPARAM)(RECT FAR*)(prc)) : (LPARAM)(RECT FAR*)NULL))


Public Const LVM_SETITEMPOSITION = (LVM_FIRST + 15)
'public const ListView_SetItemPosition(hwndLV, i, x, y) \
'    (BOOL)SNDMSG((hwndLV), LVM_SETITEMPOSITION, (WPARAM)(int)(i), MAKELPARAM((x), (y)))


Public Const LVM_GETITEMPOSITION = (LVM_FIRST + 16)
'public const ListView_GetItemPosition(hwndLV, i, ppt) \
'    (BOOL)SNDMSG((hwndLV), LVM_GETITEMPOSITION, (WPARAM)(int)(i), (LPARAM)(POINT FAR*)(ppt))


Public Const LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
Public Const LVM_GETSTRINGWIDTHW = (LVM_FIRST + 87)
#If UNICODE Then
Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHW
#Else
Public Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA
#End If

'public const ListView_GetStringWidth(hwndLV, psz) \
'    (int)SNDMSG((hwndLV), LVM_GETSTRINGWIDTH, 0, (LPARAM)(LPCTSTR)(psz))


Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Public Const LVHT_ABOVE = &H8
Public Const LVHT_BELOW = &H10
Public Const LVHT_TORIGHT = &H20
Public Const LVHT_TOLEFT = &H40

Public Type LVHITTESTINFO
    pt As POINTAPI
    flags As Long
    iItem As Long
' #if (_WIN32_IE >= =&H0300)
      iSubItem As Long ';    ' this is was NOT in win95.  valid only for LVM_SUBITEMHITTEST
' #end If
End Type

Public Const LVM_HITTEST = (LVM_FIRST + 18)
'public const ListView_HitTest(hwndLV, pinfo)
'    (int)SNDMSG((hwndLV), LVM_HITTEST, 0, (LPARAM)(LV_HITTESTINFO FAR*)(pinfo))


Public Const LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
'public const ListView_EnsureVisible(hwndLV, i, fPartialOK) \
'    (BOOL)SNDMSG((hwndLV), LVM_ENSUREVISIBLE, (WPARAM)(int)(i), MAKELPARAM((fPartialOK), 0))


Public Const LVM_SCROLL = (LVM_FIRST + 20)
'public const ListView_Scroll(hwndLV, dx, dy) \
'    (BOOL)SNDMSG((hwndLV), LVM_SCROLL, (WPARAM)(int)dx, (LPARAM)(int)dy)


Public Const LVM_REDRAWITEMS = (LVM_FIRST + 21)
'public const ListView_RedrawItems(hwndLV, iFirst, iLast) \
'    (BOOL)SNDMSG((hwndLV), LVM_REDRAWITEMS, (WPARAM)(int)iFirst, (LPARAM)(int)iLast)


Public Const LVA_DEFAULT = &H0
Public Const LVA_ALIGNLEFT = &H1
Public Const LVA_ALIGNTOP = &H2
Public Const LVA_SNAPTOGRID = &H5

Public Const LVM_ARRANGE = (LVM_FIRST + 22)
'public const ListView_Arrange(hwndLV, code) \
'    (BOOL)SNDMSG((hwndLV), LVM_ARRANGE, (WPARAM)(UINT)(code), 0L)


Public Const LVM_EDITLABELA = (LVM_FIRST + 23)
Public Const LVM_EDITLABELW = (LVM_FIRST + 118)
#If UNICODE Then
Public Const LVM_EDITLABEL = LVM_EDITLABELW
#Else
Public Const LVM_EDITLABEL = LVM_EDITLABELA
#End If

'public const ListView_EditLabel(hwndLV, i) \
'    (HWND)SNDMSG((hwndLV), LVM_EDITLABEL, (WPARAM)(int)(i), 0L)


Public Const LVM_GETEDITCONTROL = (LVM_FIRST + 24)
'public const ListView_GetEditControl(hwndLV) \
'    (HWND)SNDMSG((hwndLV), LVM_GETEDITCONTROL, 0, 0L)



#If UNICODE Then
Public Type LVCOLUMN
   mask As Long
   fmt As Long
   cx As Long
   pszText As Long
   cchTextMax As Long
   iSubItem As Long
' #if (_WIN32_IE >= =&H0300)
   iImage As Long
   iOrder As Long
' #end If
End Type
#Else
Public Type LVCOLUMN
   mask As Long
   fmt As Long
   cx As Long
   pszText As String
   cchTextMax As Long
   iSubItem As Long
' #if (_WIN32_IE >= =&H0300)
   iImage As Long
   iOrder As Long
' #end If
End Type
#End If


Public Const LVCF_FMT = &H1
Public Const LVCF_WIDTH = &H2
Public Const LVCF_TEXT = &H4
Public Const LVCF_SUBITEM = &H8
' #if (_WIN32_IE >= =&H0300)
Public Const LVCF_IMAGE = &H10
Public Const LVCF_ORDER = &H20
' #end If

Public Const LVCFMT_LEFT = &H0
Public Const LVCFMT_RIGHT = &H1
Public Const LVCFMT_CENTER = &H2
Public Const LVCFMT_JUSTIFYMASK = &H3
' #if (_WIN32_IE >= =&H0300)
Public Const LVCFMT_IMAGE = &H800
Public Const LVCFMT_BITMAP_ON_RIGHT = &H1000
Public Const LVCFMT_COL_HAS_IMAGES = &H8000
' #end If

Public Const LVM_GETCOLUMNA = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMNW = (LVM_FIRST + 95)
#If UNICODE Then
Public Const LVM_GETCOLUMN = LVM_GETCOLUMNW
#Else
Public Const LVM_GETCOLUMN = LVM_GETCOLUMNA
#End If

'public const ListView_GetColumn(hwnd, iCol, pcol) \
'    (BOOL)SNDMSG((hwnd), LVM_GETCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(LV_COLUMN FAR*)(pcol))


Public Const LVM_SETCOLUMNA = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMNW = (LVM_FIRST + 96)
#If UNICODE Then
Public Const LVM_SETCOLUMN = LVM_SETCOLUMNW
#Else
Public Const LVM_SETCOLUMN = LVM_SETCOLUMNA
#End If

'public const ListView_SetColumn(hwnd, iCol, pcol) \
'    (BOOL)SNDMSG((hwnd), LVM_SETCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(const LV_COLUMN FAR*)(pcol))


Public Const LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
Public Const LVM_INSERTCOLUMNW = (LVM_FIRST + 97)
#If UNICODE Then
Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNW
#Else
Public Const LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA
#End If

'public const ListView_InsertColumn(hwnd, iCol, pcol) \
'    (int)SNDMSG((hwnd), LVM_INSERTCOLUMN, (WPARAM)(int)(iCol), (LPARAM)(const LV_COLUMN FAR*)(pcol))


Public Const LVM_DELETECOLUMN = (LVM_FIRST + 28)
'public const ListView_DeleteColumn(hwnd, iCol) \
'    (BOOL)SNDMSG((hwnd), LVM_DELETECOLUMN, (WPARAM)(int)(iCol), 0)


Public Const LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
'public const ListView_GetColumnWidth(hwnd, iCol) \
'    (int)SNDMSG((hwnd), LVM_GETCOLUMNWIDTH, (WPARAM)(int)(iCol), 0)


Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

'public const ListView_SetColumnWidth(hwnd, iCol, cx) \
'    (BOOL)SNDMSG((hwnd), LVM_SETCOLUMNWIDTH, (WPARAM)(int)(iCol), MAKELPARAM((cx), 0))

' #if (_WIN32_IE >= =&H0300)
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
'public const ListView_GetHeader(hwnd)\
'    (HWND)SNDMSG((hwnd), LVM_GETHEADER, 0, 0L)
' #end If

Public Const LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
'public const ListView_CreateDragImage(hwnd, i, lpptUpLeft) \
'    (HIMAGELIST)SNDMSG((hwnd), LVM_CREATEDRAGIMAGE, (WPARAM)(int)(i), (LPARAM)(LPPOINT)(lpptUpLeft))


Public Const LVM_GETVIEWRECT = (LVM_FIRST + 34)
'public const ListView_GetViewRect(hwnd, prc) \
'    (BOOL)SNDMSG((hwnd), LVM_GETVIEWRECT, 0, (LPARAM)(RECT FAR*)(prc))


Public Const LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
'public const ListView_GetTextColor(hwnd)  \
'    (COLORREF)SNDMSG((hwnd), LVM_GETTEXTCOLOR, 0, 0L)


Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
'public const ListView_SetTextColor(hwnd, clrText) \
'    (BOOL)SNDMSG((hwnd), LVM_SETTEXTCOLOR, 0, (LPARAM)(COLORREF)(clrText))


Public Const LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
'public const ListView_GetTextBkColor(hwnd)  \
'    (COLORREF)SNDMSG((hwnd), LVM_GETTEXTBKCOLOR, 0, 0L)


Public Const LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
'public const ListView_SetTextBkColor(hwnd, clrTextBk) \
'    (BOOL)SNDMSG((hwnd), LVM_SETTEXTBKCOLOR, 0, (LPARAM)(COLORREF)(clrTextBk))


Public Const LVM_GETTOPINDEX = (LVM_FIRST + 39)
'public const ListView_GetTopIndex(hwndLV) \
'    (int)SNDMSG((hwndLV), LVM_GETTOPINDEX, 0, 0)


Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
'public const ListView_GetCountPerPage(hwndLV) \
'    (int)SNDMSG((hwndLV), LVM_GETCOUNTPERPAGE, 0, 0)


Public Const LVM_GETORIGIN = (LVM_FIRST + 41)
'public const ListView_GetOrigin(hwndLV, ppt) \
'    (BOOL)SNDMSG((hwndLV), LVM_GETORIGIN, (WPARAM)0, (LPARAM)(POINT FAR*)(ppt))


Public Const LVM_UPDATE = (LVM_FIRST + 42)
'public const ListView_Update(hwndLV, i) \
'    (BOOL)SNDMSG((hwndLV), LVM_UPDATE, (WPARAM)i, 0L)


Public Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
'public const ListView_SetItemState(hwndLV, i, data, mask) \
'{ LV_ITEM _ms_lvi;\
'  _ms_lvi.stateMask = mask;\
'  _ms_lvi.state = data;\
'  SNDMSG((hwndLV), LVM_SETITEMSTATE, (WPARAM)i, (LPARAM)(LV_ITEM FAR *)&_ms_lvi);\
'}

' #if (_WIN32_IE >= =&H0300)
'public const ListView_SetCheckState(hwndLV, i, fCheck) \
'  ListView_SetItemState(hwndLV, i, INDEXTOSTATEIMAGEMASK((fCheck)?2:1), LVIS_STATEIMAGEMASK)
' #end If

Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
'public const ListView_GetItemState(hwndLV, i, mask) \
'   (UINT)SNDMSG((hwndLV), LVM_GETITEMSTATE, (WPARAM)i, (LPARAM)mask)

' #if (_WIN32_IE >= =&H0300)
'public const ListView_GetCheckState(hwndLV, i) \
'   ((((UINT)(SNDMSG((hwndLV), LVM_GETITEMSTATE, (WPARAM)i, LVIS_STATEIMAGEMASK))) >> 12) -1)
' #end If

Public Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
Public Const LVM_GETITEMTEXTW = (LVM_FIRST + 115)

#If UNICODE Then
Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTW
#Else
Public Const LVM_GETITEMTEXT = LVM_GETITEMTEXTA
#End If

'public const ListView_GetItemText(hwndLV, i, iSubItem_, pszText_, cchTextMax_) \
'{ LV_ITEM _ms_lvi;\
'  _ms_lvi.iSubItem = iSubItem_;\
'  _ms_lvi.cchTextMax = cchTextMax_;\
'  _ms_lvi.pszText = pszText_;\
'  SNDMSG((hwndLV), LVM_GETITEMTEXT, (WPARAM)i, (LPARAM)(LV_ITEM FAR *)&_ms_lvi);\
'}


Public Const LVM_SETITEMTEXTA = (LVM_FIRST + 46)
Public Const LVM_SETITEMTEXTW = (LVM_FIRST + 116)

#If UNICODE Then
Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTW
#Else
Public Const LVM_SETITEMTEXT = LVM_SETITEMTEXTA
#End If

'public const ListView_SetItemText(hwndLV, i, iSubItem_, pszText_) \
'{ LV_ITEM _ms_lvi;\
'  _ms_lvi.iSubItem = iSubItem_;\
'  _ms_lvi.pszText = pszText_;\
'  SNDMSG((hwndLV), LVM_SETITEMTEXT, (WPARAM)i, (LPARAM)(LV_ITEM FAR *)&_ms_lvi);\
'}

' #if (_WIN32_IE >= =&H0300)
' these flags only apply to LVS_OWNERDATA listviews in report or list mode
Public Const LVSICF_NOINVALIDATEALL = &H1
Public Const LVSICF_NOSCROLL = &H2
' #end If

Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
'public const ListView_SetItemCount(hwndLV, cItems) \
'  SNDMSG((hwndLV), LVM_SETITEMCOUNT, (WPARAM)cItems, 0)

' #if (_WIN32_IE >= =&H0300)
'public const ListView_SetItemCountEx(hwndLV, cItems, dwFlags) \
'  SNDMSG((hwndLV), LVM_SETITEMCOUNT, (WPARAM)cItems, (LPARAM)dwFlags)
' #end If

'public typedef int (CALLBACK *PFNLVCOMPARE)(LPARAM, LPARAM, LPARAM);


Public Const LVM_SORTITEMS = (LVM_FIRST + 48)
'public const ListView_SortItems(hwndLV, _pfnCompare, _lPrm) \
'  (BOOL)SNDMSG((hwndLV), LVM_SORTITEMS, (WPARAM)(LPARAM)_lPrm, \
'  (LPARAM)(PFNLVCOMPARE)_pfnCompare)


Public Const LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
'public const ListView_SetItemPosition32(hwndLV, i, x0, y0) \
'{   POINT ptNewPos; \
'    ptNewPos.x = x0; ptNewPos.y = y0; \
'    SNDMSG((hwndLV), LVM_SETITEMPOSITION32, (WPARAM)(int)(i), (LPARAM)&ptNewPos); \
'}


Public Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
'public const ListView_GetSelectedCount(hwndLV) \
'    (UINT)SNDMSG((hwndLV), LVM_GETSELECTEDCOUNT, 0, 0L)


Public Const LVM_GETITEMSPACING = (LVM_FIRST + 51)
'public const ListView_GetItemSpacing(hwndLV, fSmall) \
'        (DWORD)SNDMSG((hwndLV), LVM_GETITEMSPACING, fSmall, 0L)


Public Const LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
Public Const LVM_GETISEARCHSTRINGW = (LVM_FIRST + 117)

#If UNICODE Then
Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGW
#Else
Public Const LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA
#End If

'public const ListView_GetISearchString(hwndLV, lpsz) \
'        (BOOL)SNDMSG((hwndLV), LVM_GETISEARCHSTRING, 0, (LPARAM)(LPTSTR)lpsz)

' #if (_WIN32_IE >= =&H0300)
Public Const LVM_SETICONSPACING = (LVM_FIRST + 53)
' -1 for cx and cy means we'll use the default (system settings)
' 0 for cx or cy means use the current setting (allows you to change just one param)
'public const ListView_SetIconSpacing(hwndLV, cx, cy) \
'        (DWORD)SNDMSG((hwndLV), LVM_SETICONSPACING, 0, MAKELONG(cx,cy))


Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)   ' optional wParam == mask
'public const ListView_SetExtendedListViewStyle(hwndLV, dw)\
'        (DWORD)SNDMSG((hwndLV), LVM_SETEXTENDEDLISTVIEWSTYLE, 0, dw)
' #if (_WIN32_IE >= =&H0400)
'public const ListView_SetExtendedListViewStyleEx(hwndLV, dwMask, dw)\
'        (DWORD)SNDMSG((hwndLV), LVM_SETEXTENDEDLISTVIEWSTYLE, dwMask, dw)
' #end If

Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
'public const ListView_GetExtendedListViewStyle(hwndLV)\
'        (DWORD)SNDMSG((hwndLV), LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)

Public Const LVS_EX_GRIDLINES = &H1&
Public Const LVS_EX_SUBITEMIMAGES = &H2&
Public Const LVS_EX_CHECKBOXES = &H4&
Public Const LVS_EX_TRACKSELECT = &H8&
Public Const LVS_EX_HEADERDRAGDROP = &H10&
Public Const LVS_EX_FULLROWSELECT = &H20&         ' applies to report mode only
Public Const LVS_EX_ONECLICKACTIVATE = &H40&
Public Const LVS_EX_TWOCLICKACTIVATE = &H80&
' #if (_WIN32_IE >= =&H0400)
Public Const LVS_EX_FLATSB = &H100&
Public Const LVS_EX_REGIONAL = &H200&
Public Const LVS_EX_INFOTIP = &H400&              ' listview does InfoTips for you
Public Const LVS_EX_UNDERLINEHOT = &H800&
Public Const LVS_EX_UNDERLINECOLD = &H1000&
Public Const LVS_EX_MULTIWORKAREAS = &H2000&
' #end If
'#if (_WIN32_IE >= 0x0500)
Public Const LVS_EX_LABELTIP = &H4000&             '// listview unfolds partly hidden labels if it does not have infotip text
Public Const LVS_EX_BORDERSELECT = &H8000&         '// border selection style instead of highlight
'#endif  // End (_WIN32_IE >= = &H0500)
'#if (_WIN32_WINNT >= 0x501)
Public Const LVS_EX_DOUBLEBUFFER = &H10000
Public Const LVS_EX_HIDELABELS = &H20000
Public Const LVS_EX_SINGLEROW = &H40000
Public Const LVS_EX_SNAPTOGRID = &H80000           '// Icons automatically snap to grid.
Public Const LVS_EX_SIMPLESELECT = &H100000         '// Also changes overlay rendering to top right for icon mode.
'#End If


Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
'public const ListView_GetSubItemRect(hwnd, iItem, iSubItem, code, prc) \
'        (BOOL)SNDMSG((hwnd), LVM_GETSUBITEMRECT, (WPARAM)(int)(iItem), \
'                ((prc) ? ((((LPRECT)(prc))->top = iSubItem), (((LPRECT)(prc))->left = code), (LPARAM)(prc)) : (LPARAM)(LPRECT)NULL))

Public Const LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
'public const ListView_SubItemHitTest(hwnd, plvhti) \
'        (int)SNDMSG((hwnd), LVM_SUBITEMHITTEST, 0, (LPARAM)(LPLVHITTESTINFO)(plvhti))

Public Const LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
'public const ListView_SetColumnOrderArray(hwnd, iCount, pi) \
'        (BOOL)SNDMSG((hwnd), LVM_SETCOLUMNORDERARRAY, (WPARAM)iCount, (LPARAM)(LPINT)pi)

Public Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
'public const ListView_GetColumnOrderArray(hwnd, iCount, pi) \
'        (BOOL)SNDMSG((hwnd), LVM_GETCOLUMNORDERARRAY, (WPARAM)iCount, (LPARAM)(LPINT)pi)

Public Const LVM_SETHOTITEM = (LVM_FIRST + 60)
'public const ListView_SetHotItem(hwnd, i) \
'        (int)SNDMSG((hwnd), LVM_SETHOTITEM, (WPARAM)i, 0)

Public Const LVM_GETHOTITEM = (LVM_FIRST + 61)
'public const ListView_GetHotItem(hwnd) \
'        (int)SNDMSG((hwnd), LVM_GETHOTITEM, 0, 0)

Public Const LVM_SETHOTCURSOR = (LVM_FIRST + 62)
'public const ListView_SetHotCursor(hwnd, hcur) \
'        (HCURSOR)SNDMSG((hwnd), LVM_SETHOTCURSOR, 0, (LPARAM)hcur)

Public Const LVM_GETHOTCURSOR = (LVM_FIRST + 63)
'public const ListView_GetHotCursor(hwnd) \
'        (HCURSOR)SNDMSG((hwnd), LVM_GETHOTCURSOR, 0, 0)

Public Const LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
'public const ListView_ApproximateViewRect(hwnd, iWidth, iHeight, iCount) \
'        (DWORD)SNDMSG((hwnd), LVM_APPROXIMATEVIEWRECT, iCount, MAKELPARAM(iWidth, iHeight))
' #end If     ' _WIN32_IE >= =&H0300

' #if (_WIN32_IE >= =&H0400)

Public Const LV_MAX_WORKAREAS = 16
Public Const LVM_SETWORKAREAS = (LVM_FIRST + 65)
'public const ListView_SetWorkAreas(hwnd, nWorkAreas, prc) \
'    (BOOL)SNDMSG((hwnd), LVM_SETWORKAREAS, (WPARAM)(int)nWorkAreas, (LPARAM)(RECT FAR*)(prc))

Public Const LVM_GETWORKAREAS = (LVM_FIRST + 70)
'public const ListView_GetWorkAreas(hwnd, nWorkAreas, prc) \
'    (BOOL)SNDMSG((hwnd), LVM_GETWORKAREAS, (WPARAM)(int)nWorkAreas, (LPARAM)(RECT FAR*)(prc))


Public Const LVM_GETNUMBEROFWORKAREAS = (LVM_FIRST + 73)
'public const ListView_GetNumberOfWorkAreas(hwnd, pnWorkAreas) \
'    (BOOL)SNDMSG((hwnd), LVM_GETNUMBEROFWORKAREAS, 0, (LPARAM)(UINT *)(pnWorkAreas))


Public Const LVM_GETSELECTIONMARK = (LVM_FIRST + 66)
'public const ListView_GetSelectionMark(hwnd) \
'    (int)SNDMSG((hwnd), LVM_GETSELECTIONMARK, 0, 0)

Public Const LVM_SETSELECTIONMARK = (LVM_FIRST + 67)
'public const ListView_SetSelectionMark(hwnd, i) \
'    (int)SNDMSG((hwnd), LVM_SETSELECTIONMARK, 0, (LPARAM)i)

Public Const LVM_SETHOVERTIME = (LVM_FIRST + 71)
'public const ListView_SetHoverTime(hwndLV, dwHoverTimeMs)\
'        (DWORD)SendMessage((hwndLV), LVM_SETHOVERTIME, 0, dwHoverTimeMs)

Public Const LVM_GETHOVERTIME = (LVM_FIRST + 72)
'public const ListView_GetHoverTime(hwndLV)\
'        (DWORD)SendMessage((hwndLV), LVM_GETHOVERTIME, 0, 0)

Public Const LVM_SETTOOLTIPS = (LVM_FIRST + 74)
'public const ListView_SetToolTips(hwndLV, hwndNewHwnd)\
'        (HWND)SendMessage((hwndLV), LVM_SETTOOLTIPS, hwndNewHwnd, 0)

Public Const LVM_GETTOOLTIPS = (LVM_FIRST + 78)
'public const ListView_GetToolTips(hwndLV)\
'        (HWND)SendMessage((hwndLV), LVM_GETTOOLTIPS, 0, 0)


Public Const LVM_SORTITEMSEX = (LVM_FIRST + 81)
'public const ListView_SortItemsEx(hwndLV, _pfnCompare, _lPrm) \
'  (BOOL)SNDMSG((hwndLV), LVM_SORTITEMSEX, (WPARAM)(LPARAM)_lPrm, (LPARAM)(PFNLVCOMPARE)_pfnCompare)

#If UNICODE Then
Public Type LVBKIMAGE
    ulFlags As Long ';              ' LVBKIF_*
    hbm As Long
    pszImage As Long
    cchImageMax As Long
    xOffsetPercent As Long
    yOffsetPercent As Long
End Type
#Else
Public Type LVBKIMAGE
    ulFlags As Long ';              ' LVBKIF_*
    hbm As Long
    pszImage As String
    cchImageMax As Long
    xOffsetPercent As Long
    yOffsetPercent As Long
End Type
#End If

Public Const LVBKIF_SOURCE_NONE = &H0
Public Const LVBKIF_SOURCE_HBITMAP = &H1
Public Const LVBKIF_SOURCE_URL = &H2
Public Const LVBKIF_SOURCE_MASK = &H3
Public Const LVBKIF_STYLE_NORMAL = &H0
Public Const LVBKIF_STYLE_TILE = &H10
Public Const LVBKIF_STYLE_MASK = &H10

Public Const LVM_SETBKIMAGEA = (LVM_FIRST + 68)
Public Const LVM_SETBKIMAGEW = (LVM_FIRST + 138)
Public Const LVM_GETBKIMAGEA = (LVM_FIRST + 69)
Public Const LVM_GETBKIMAGEW = (LVM_FIRST + 139)

#If UNICODE Then
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEW
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEW
#Else
Public Const LVM_SETBKIMAGE = LVM_SETBKIMAGEA
Public Const LVM_GETBKIMAGE = LVM_GETBKIMAGEA
#End If

'public const ListView_SetBkImage(hwnd, plvbki) \
'    (BOOL)SNDMSG((hwnd), LVM_SETBKIMAGE, 0, (LPARAM)plvbki)

'public const ListView_GetBkImage(hwnd, plvbki) \
'    (BOOL)SNDMSG((hwnd), LVM_GETBKIMAGE, 0, (LPARAM)plvbki)


' #end If     ' _WIN32_IE >= =&H0400

' #end If

'#if (_WIN32_WINNT >= = &H501)
Public Const LVM_SETSELECTEDCOLUMN = (LVM_FIRST + 140)
'public const ListView_SetSelectedColumn(hwnd, iCol) \
'    SNDMSG((hwnd), LVM_SETSELECTEDCOLUMN, (WPARAM)iCol, 0)

Public Const LVM_SETTILEWIDTH = (LVM_FIRST + 141)
'public const ListView_SetTileWidth(hwnd, cpWidth) \
'    SNDMSG((hwnd), LVM_SETTILEWIDTH, (WPARAM)cpWidth, 0)

Public Const LV_VIEW_ICON = &H0&        '= &H0000
Public Const LV_VIEW_DETAILS = &H1&      '= &H0001
Public Const LV_VIEW_SMALLICON = &H2&   '= &H0002
Public Const LV_VIEW_LIST = &H3&        '= &H0003
Public Const LV_VIEW_TILE = &H4&        '= &H0004

Public Const LVM_SETVIEW = (LVM_FIRST + 142)
'public const ListView_SetView(hwnd, iView) \
'    (DWORD)SNDMSG((hwnd), LVM_SETVIEW, (WPARAM)(DWORD)iView, 0)

Public Const LVM_GETVIEW = (LVM_FIRST + 143)
'public const ListView_GetView(hwnd) \
'    (DWORD)SNDMSG((hwnd), LVM_GETVIEW, 0, 0)


Public Const LVGF_NONE = &H0&
Public Const LVGF_HEADER = &H1&
Public Const LVGF_FOOTER = &H2&
Public Const LVGF_STATE = &H4&
Public Const LVGF_ALIGN = &H8&
Public Const LVGF_GROUPID = &H10&

Public Const LVGS_NORMAL = &H0&
Public Const LVGS_COLLAPSED = &H1&
Public Const LVGS_HIDDEN = &H2&

Public Const LVGA_HEADER_LEFT = &H1&
Public Const LVGA_HEADER_CENTER = &H2&
Public Const LVGA_HEADER_RIGHT = &H4&  '// Don't forget to validate exclusivity
Public Const LVGA_FOOTER_LEFT = &H8&
Public Const LVGA_FOOTER_CENTER = &H10&
Public Const LVGA_FOOTER_RIGHT = &H20&   '// Don't forget to validate exclusivity

Public Type LVGROUP 'typedef struct tagLVGROUP
    cbSize As Long
    mask As Long
    
    pszHeader As Long
    cchHeader As Long

    pszFooter As Long
    cchFooter As Long

    iGroupId As Long

    stateMask As Long
    state As Long
    uAlign As Long
End Type

Public Const LVM_INSERTGROUP = (LVM_FIRST + 145)
'public const ListView_InsertGroup(hwnd, index, pgrp) \
'    SNDMSG((hwnd), LVM_INSERTGROUP, (WPARAM)index, (LPARAM)pgrp)


Public Const LVM_SETGROUPINFO = (LVM_FIRST + 147)
'public const ListView_SetGroupInfo(hwnd, iGroupId, pgrp) \
'    SNDMSG((hwnd), LVM_SETGROUPINFO, (WPARAM)iGroupId, (LPARAM)pgrp)


Public Const LVM_GETGROUPINFO = (LVM_FIRST + 149)
'public const ListView_GetGroupInfo(hwnd, iGroupId, pgrp) \
'    SNDMSG((hwnd), LVM_GETGROUPINFO, (WPARAM)iGroupId, (LPARAM)pgrp)


Public Const LVM_REMOVEGROUP = (LVM_FIRST + 150)
'public const ListView_RemoveGroup(hwnd, iGroupId) \
'    SNDMSG((hwnd), LVM_REMOVEGROUP, (WPARAM)iGroupId, 0)

Public Const LVM_MOVEGROUP = (LVM_FIRST + 151)
'public const ListView_MoveGroup(hwnd, iGroupId, toIndex) \
'    SNDMSG((hwnd), LVM_MOVEGROUP, (WPARAM)iGroupId, (LPARAM)toIndex)

Public Const LVM_MOVEITEMTOGROUP = (LVM_FIRST + 154)
'public const ListView_MoveItemToGroup(hwnd, idItemFrom, idGroupTo) \
'    SNDMSG((hwnd), LVM_MOVEITEMTOGROUP, (WPARAM)idItemFrom, (LPARAM)idGroupTo)


Public Const LVGMF_NONE = &H0&
Public Const LVGMF_BORDERSIZE = &H1&
Public Const LVGMF_BORDERCOLOR = &H2&
Public Const LVGMF_TEXTCOLOR = &H4&

Public Type LVGROUPMETRICS 'struct tagLVGROUPMETRICS
    cbSize As Long
    mask As Long
    left As Long
    top As Long
    right As Long
    bottom As Long
    crLeft As Long
    crTop As Long
    crRight As Long
    crBottom As Long
    crHeader As Long
    crFooter As Long
End Type

Public Const LVM_SETGROUPMETRICS = (LVM_FIRST + 155)
'public const ListView_SetGroupMetrics(hwnd, pGroupMetrics) \
'    SNDMSG((hwnd), LVM_SETGROUPMETRICS, 0, (LPARAM)pGroupMetrics)

Public Const LVM_GETGROUPMETRICS = (LVM_FIRST + 156)
'public const ListView_GetGroupMetrics(hwnd, pGroupMetrics) \
'    SNDMSG((hwnd), LVM_GETGROUPMETRICS, 0, (LPARAM)pGroupMetrics)

Public Const LVM_ENABLEGROUPVIEW = (LVM_FIRST + 157)
'public const ListView_EnableGroupView(hwnd, fEnable) \
'    SNDMSG((hwnd), LVM_ENABLEGROUPVIEW, (WPARAM)fEnable, 0)
'
'typedef int (CALLBACK *PFNLVGROUPCOMPARE)(int, int, void *);

Public Const LVM_SORTGROUPS = (LVM_FIRST + 158)
'public const ListView_SortGroups(hwnd, _pfnGroupCompate, _plv) \
'    SNDMSG((hwnd), LVM_SORTGROUPS, (WPARAM)_pfnGroupCompate, (LPARAM)_plv)

Public Type LVINSERTGROUPSORTED
    pfnGroupCompare As Long
    pvData As Long
    LVGROUP As LVGROUP
End Type

Public Const LVM_INSERTGROUPSORTED = (LVM_FIRST + 159)
'public const ListView_InsertGroupSorted(hwnd, structInsert) \
'    SNDMSG((hwnd), LVM_INSERTGROUPSORTED, (WPARAM)structInsert, 0)

Public Const LVM_REMOVEALLGROUPS = (LVM_FIRST + 160)
'public const ListView_RemoveAllGroups(hwnd) \
'    SNDMSG((hwnd), LVM_REMOVEALLGROUPS, 0, 0)

Public Const LVM_HASGROUP = (LVM_FIRST + 161)
'public const ListView_HasGroup(hwnd, dwGroupId) \
'    SNDMSG((hwnd), LVM_HASGROUP, dwGroupId, 0)


Public Const LVTVIF_AUTOSIZE = &H0
Public Const LVTVIF_FIXEDWIDTH = &H1
Public Const LVTVIF_FIXEDHEIGHT = &H2
Public Const LVTVIF_FIXEDSIZE = &H3

Public Const LVTVIM_TILESIZE = &H1
Public Const LVTVIM_COLUMNS = &H2
Public Const LVTVIM_LABELMARGIN = &H4

Public Type LVTILEVIEWINFO
    cbSize As Long
    dwMask As Long ';     //LVTVIM_*
    dwFlags As Long ';    //LVTVIF_*
    sizeTile As Size ' ;
    cLines As Long
    rcLabelMargin As RECT
End Type

Public Type LVTILEINFO
    cbSize As Long
    iItem As Long
    cColumns As Long
    puColumns As Long
End Type

Public Const LVM_SETTILEVIEWINFO = (LVM_FIRST + 162)
'public const ListView_SetTileViewInfo(hwnd, ptvi) \
'    SNDMSG((hwnd), LVM_SETTILEVIEWINFO, 0, (LPARAM)ptvi)

Public Const LVM_GETTILEVIEWINFO = (LVM_FIRST + 163)
'public const ListView_GetTileViewInfo(hwnd, ptvi) \
'    SNDMSG((hwnd), LVM_GETTILEVIEWINFO, 0, (LPARAM)ptvi)

Public Const LVM_SETTILEINFO = (LVM_FIRST + 164)
'public const ListView_SetTileInfo(hwnd, pti) \
'    SNDMSG((hwnd), LVM_SETTILEINFO, 0, (LPARAM)pti)

Public Const LVM_GETTILEINFO = (LVM_FIRST + 165)
'public const ListView_GetTileInfo(hwnd, pti) \
'    SNDMSG((hwnd), LVM_GETTILEINFO, 0, (LPARAM)pti)

Public Type LVINSERTMARK
    cbSize As Long
    dwFlags As Long
    iItem As Long
    dwReserved As Long
End Type

Public Const LVIM_AFTER = &H1&              '// TRUE = insert After iItem, otherwise before

Public Const LVM_SETINSERTMARK = (LVM_FIRST + 166)
'public const ListView_SetInsertMark(hwnd, lvim) \
'    (BOOL)SNDMSG((hwnd), LVM_SETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))

Public Const LVM_GETINSERTMARK = (LVM_FIRST + 167)
'public const ListView_GetInsertMark(hwnd, lvim) \
'    (BOOL)SNDMSG((hwnd), LVM_GETINSERTMARK, (WPARAM) 0, (LPARAM) (lvim))

Public Const LVM_INSERTMARKHITTEST = (LVM_FIRST + 168)
'public const ListView_InsertMarkHitTest(hwnd, point, lvim) \
'    (int)SNDMSG((hwnd), LVM_INSERTMARKHITTEST, (WPARAM)(LPPOINT)(point), (LPARAM)(LPLVINSERTMARK)(lvim))

Public Const LVM_GETINSERTMARKRECT = (LVM_FIRST + 169)
'public const ListView_GetInsertMarkRect(hwnd, rc) \
'    (int)SNDMSG((hwnd), LVM_GETINSERTMARKRECT, (WPARAM)0, (LPARAM)(LPRECT)(rc))

Public Const LVM_SETINSERTMARKCOLOR = (LVM_FIRST + 170)
'public const ListView_SetInsertMarkColor(hwnd, color) \
'    (COLORREF)SNDMSG((hwnd), LVM_SETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))

Public Const LVM_GETINSERTMARKCOLOR = (LVM_FIRST + 171)
'public const ListView_GetInsertMarkColor(hwnd) \
'    (COLORREF)SNDMSG((hwnd), LVM_GETINSERTMARKCOLOR, (WPARAM)0, (LPARAM)0)

Public Type LVSETINFOTIP
    cbSize As Long
    dwFlags As Long
    pszText As Long ' LPWSTR
    iItem As Long
    iSubItem As Long
End Type

Public Const LVM_SETINFOTIP = (LVM_FIRST + 173)

'public const ListView_SetInfoTip(hwndLV, plvInfoTip)\
'        (BOOL)SNDMSG((hwndLV), LVM_SETINFOTIP, (WPARAM)0, (LPARAM)plvInfoTip)

Public Const LVM_GETSELECTEDCOLUMN = (LVM_FIRST + 174)
'public const ListView_GetSelectedColumn(hwnd) \
'    (UINT)SNDMSG((hwnd), LVM_GETSELECTEDCOLUMN, 0, 0)


Public Const LVM_ISGROUPVIEWENABLED = (LVM_FIRST + 175)
'public const ListView_IsGroupViewEnabled(hwnd) \
'    (BOOL)SNDMSG((hwnd), LVM_ISGROUPVIEWENABLED, 0, 0)

Public Const LVM_GETOUTLINECOLOR = (LVM_FIRST + 176)
'public const ListView_GetOutlineColor(hwnd) \
'    (COLORREF)SNDMSG((hwnd), LVM_GETOUTLINECOLOR, 0, 0)

Public Const LVM_SETOUTLINECOLOR = (LVM_FIRST + 177)
'public const ListView_SetOutlineColor(hwnd, color) \
'    (COLORREF)SNDMSG((hwnd), LVM_SETOUTLINECOLOR, (WPARAM)0, (LPARAM)(COLORREF)(color))


Public Const LVM_CANCELEDITLABEL = (LVM_FIRST + 179)
'public const ListView_CancelEditLabel(hwnd) \
'    (VOID)SNDMSG((hwnd), LVM_CANCELEDITLABEL, (WPARAM)0, (LPARAM)0)


'// These next to methods make it easy to identify an item that can be repositioned
'// within listview. For example: Many developers use the lParam to store an identifier that is
'// unique. Unfortunatly, in order to find this item, they have to iterate through all of the items
'// in the listview. Listview will maintain a unique identifier.  The upper bound is the size of a DWORD.
Public Const LVM_MAPINDEXTOID = (LVM_FIRST + 180)
'public const ListView_MapIndexToID(hwnd, index) \
'    (UINT)SNDMSG((hwnd), LVM_MAPINDEXTOID, (WPARAM)index, (LPARAM)0)

Public Const LVM_MAPIDTOINDEX = (LVM_FIRST + 181)
'public const ListView_MapIDToIndex(hwnd, id) \
'    (UINT)SNDMSG((hwnd), LVM_MAPIDTOINDEX, (WPARAM)id, (LPARAM)0)

'#End If


Public Type NMLISTVIEW
    hdr As NMHDR
    iItem As Long
    iSubItem As Long
    uNewState As Long
    uOldState As Long
    uChanged As Long
    ptAction As POINTAPI
    lParam As Long
End Type


' #if (_WIN32_IE >= =&H400)
' NMITEMACTIVATE is used instead of NMLISTVIEW in IE >= =&H400
' therefore all the fields are the same except for extra uKeyFlags
' they are used to store key flags at the time of the single click with
' delayed activation - because by the time the timer goes off a user may
' not hold the keys (shift, ctrl) any more
Public Type NMITEMACTIVATE
    hdr As NMHDR
    iItem As Long
    iSubItem As Long
    uNewState As Long
    uOldState As Long
    uChanged As Long
    ptAction As POINTAPI
    lParam As Long
    uKeyFlags As Long
End Type

' key flags stored in uKeyFlags
Public Const LVKF_ALT = &H1
Public Const LVKF_CONTROL = &H2
Public Const LVKF_SHIFT = &H4
' #end If '(_WIN32_IE >= =&H0400)


' #if (_WIN32_IE >= =&H0300)
'public const NMLVCUSTOMDRAW_V3_SIZE CCSIZEOF_STRUCT(NMLVCUSTOMDRW, clrTextBk)

Public Type NMLVCUSTOMDRAW
    nmcd As NMCUSTOMDRAW
    clrText As Long
    clrTextBk As Long
' #if (_WIN32_IE >= =&H0400)
    iSubItem As Long
' #end If
End Type

Public Type NMLVCACHEHINT
    hdr As NMHDR
    iFrom As Long
    iTo As Long
End Type

Public Type NMLVFINDITEM
    hdr As NMHDR
    iStart As Long
    lvfi As LVFINDINFO
End Type

Public Type NMLVODSTATECHANGE
    hdr As NMHDR
    iFrom As Long
    iTo As Long
    uNewState As Long
    uOldState As Long
End Type

' #end If     ' _WIN32_IE >= =&H0300

Public Const LVN_ITEMCHANGING = (LVN_FIRST - 0)
Public Const LVN_ITEMCHANGED = (LVN_FIRST - 1)
Public Const LVN_INSERTITEM = (LVN_FIRST - 2)
Public Const LVN_DELETEITEM = (LVN_FIRST - 3)
Public Const LVN_DELETEALLITEMS = (LVN_FIRST - 4)
Public Const LVN_BEGINLABELEDITA = (LVN_FIRST - 5)
Public Const LVN_BEGINLABELEDITW = (LVN_FIRST - 75)
Public Const LVN_ENDLABELEDITA = (LVN_FIRST - 6)
Public Const LVN_ENDLABELEDITW = (LVN_FIRST - 76)
Public Const LVN_COLUMNCLICK = (LVN_FIRST - 8)
Public Const LVN_BEGINDRAG = (LVN_FIRST - 9)
Public Const LVN_BEGINRDRAG = (LVN_FIRST - 11)

' #if (_WIN32_IE >= =&H0300)
Public Const LVN_ODCACHEHINT = (LVN_FIRST - 13)
Public Const LVN_ODFINDITEMA = (LVN_FIRST - 52)
Public Const LVN_ODFINDITEMW = (LVN_FIRST - 79)

Public Const LVN_ITEMACTIVATE = (LVN_FIRST - 14)
Public Const LVN_ODSTATECHANGED = (LVN_FIRST - 15)

#If UNICODE Then
Public Const LVN_ODFINDITEM = LVN_ODFINDITEMW
#Else
Public Const LVN_ODFINDITEM = LVN_ODFINDITEMA
#End If
' #end If     ' _WIN32_IE >= =&H0300


' #if (_WIN32_IE >= =&H0400)
Public Const LVN_HOTTRACK = (LVN_FIRST - 21)
' #end If

Public Const LVN_GETDISPINFOA = (LVN_FIRST - 50)
Public Const LVN_GETDISPINFOW = (LVN_FIRST - 77)
Public Const LVN_SETDISPINFOA = (LVN_FIRST - 51)
Public Const LVN_SETDISPINFOW = (LVN_FIRST - 78)

#If UNICODE Then
Public Const LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITW
Public Const LVN_ENDLABELEDIT = LVN_ENDLABELEDITW
Public Const LVN_GETDISPINFO = LVN_GETDISPINFOW
Public Const LVN_SETDISPINFO = LVN_SETDISPINFOW
#Else
Public Const LVN_BEGINLABELEDIT = LVN_BEGINLABELEDITA
Public Const LVN_ENDLABELEDIT = LVN_ENDLABELEDITA
Public Const LVN_GETDISPINFO = LVN_GETDISPINFOA
Public Const LVN_SETDISPINFO = LVN_SETDISPINFOA
#End If

Public Const LVIF_DI_SETITEM = &H1000

Public Type NMLVDISPINFO
    hdr As NMHDR
    Item As LVITEM_LT
End Type

Public Const LVN_KEYDOWN = (LVN_FIRST - 55)


Public Type NMLVKEYDOWN
    hdr As NMHDR
    wVKey As Integer
    flags1 As Integer
    flags2 As Integer
    'UINT flags;
End Type

' #if (_WIN32_IE >= =&H0300)
Public Const LVN_MARQUEEBEGIN = (LVN_FIRST - 56)
' #end If

' #if (_WIN32_IE >= =&H0400)
#If UNICODE Then
Public Type NMLVGETINFOTIP
    hdr As NMHDR
    dwFlags As Long
    pszText As Long
    cchTextMax As Long
    iItem As Long
    iSubItem As Long
    lParam As Long
End Type
#Else
Public Type NMLVGETINFOTIP
    hdr As NMHDR
    dwFlags As Long
    pszText As String
    cchTextMax As Long
    iItem As Long
    iSubItem As Long
    lParam As Long
End Type
#End If
Public Type NMLVGETINFOTIP_NOSTRING
   hdr As NMHDR
   dwFlags As Long
   pszText As Long
   cchTextMax As Long
   iItem As Long
   iSubItem As Long
   lParam As Long
End Type

Public Const LVGIT_UNFOLDED = &H1

Public Const LVN_GETINFOTIPA = (LVN_FIRST - 57)
Public Const LVN_GETINFOTIPW = (LVN_FIRST - 58)

#If UNICODE Then
Public Const LVN_GETINFOTIP = LVN_GETINFOTIPW
#Else
Public Const LVN_GETINFOTIP = LVN_GETINFOTIPA
#End If


' #end If     ' _WIN32_IE >= =&H0400

' #end If ' NOLISTVIEW

' =========================================================================================================================================

