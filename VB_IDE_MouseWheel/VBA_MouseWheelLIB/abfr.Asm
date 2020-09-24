   .586                      ; force 32 bit code
      .model flat, stdcall      ; memory model & calling convention
      option casemap :none      ; case sensitive
      .k3d
include abfr.inc

HookP PROTO :DWORD,:DWORD,:DWORD
NewProc PROTO :DWORD,:DWORD,:DWORD,:DWORD
EnumChilds PROTO :DWORD,:DWORD
EnumTop PROTO :DWORD,:DWORD
EnumChildsFree PROTO :DWORD,:DWORD
TestWindow PROTO :DWORD,:DWORD
SetLineNumber PROTO :DWORD
ScrollIt PROTO :DWORD,:DWORD,:DWORD
;MOUSEWHEEL FOR VB6 by Vanja Fuckar!!!!!!!!!!

.data

GlobalHook dd 0
AppInst dd 0

OP db "OldProc",0
SBAR db "SBar",0
 
VbaWindow db "VbaWindow",0
ScrollBar db "ScrollBar",0
VB6MV db "VB6IDEMOUSEWHEEL",0
ScrollLines db  "Lines",0


 IsAttach db 0
 IsDetach db 0
 
.code



DllEntry proc hInstance:DWORD, reason:DWORD, reserved1:DWORD
mov eax,reason


.if eax==DLL_PROCESS_DETACH && IsDetach==0
invoke EnumWindows,EnumTop,0 ;free resources
or IsDetach,1

.elseif eax==DLL_PROCESS_ATTACH && IsAttach==0
invoke EnumWindows,EnumTop,1 ;join resources
or IsAttach,1
mov eax,hInstance
mov AppInst,eax

.endif

mov  eax,1
ret
DllEntry Endp


EnableWheel proc
invoke SetWindowsHookEx,5,HookP,AppInst,0
mov GlobalHook,eax
ret
EnableWheel endp

DisableWheel proc
invoke UnhookWindowsHookEx,GlobalHook
ret
DisableWheel endp

TestWindow proc public uses esi hWnd:DWORD,ClassName:DWORD
sub esp,256
mov esi,esp
invoke GetClassName,hWnd,esi,256
invoke lstrcmpi,esi,ClassName
add esp,256
ret
TestWindow endp

HookP proc public uses esi edi ebx lMsg:DWORD,wParam:DWORD,lParam:DWORD

invoke TestWindow,wParam,addr VbaWindow

mov esi,eax


	.if lMsg==HCBT_CREATEWND
	
		test esi,esi ;is VBA WINDOW
		jne @f
			invoke SetWindowLong,wParam,GWL_WNDPROC,NewProc
			invoke SetProp,wParam,addr OP,eax
			invoke SetLineNumber,wParam		
			
		@@:
			
	.elseif lMsg==HCBT_DESTROYWND
		
		test esi,esi ;is VBA WINDOW
		jne @f
			invoke GetProp,wParam,addr OP
			invoke SetWindowLong,wParam,GWL_WNDPROC,eax
			invoke RemoveProp,wParam,addr OP
			invoke RemoveProp,wParam,addr ScrollLines	
		@@:
	
	.endif
		

invoke CallNextHookEx,GlobalHook,lMsg,wParam,lParam
ret
HookP endp


SetLineNumber proc hWnd:DWORD
	sub esp,4
	invoke RegOpenKey,080000000h,addr VB6MV,esp
	pop ebx
			
	test eax,eax
	jne @f ;nema
	sub esp,12
	mov ecx,esp
	lea edx,[esp+4]
	lea eax,[esp+8]
	invoke RegQueryValueEx,ebx,addr ScrollLines,0,ecx,edx,eax
	invoke RegCloseKey,ebx
	mov edx,DWORD PTR [esp+4] ;uzmi value
	add esp,12	
	test eax,eax
	jne @f ;nema
	invoke SetProp,hWnd,addr ScrollLines,edx			
	@@:

ret
SetLineNumber endp

NewProc proc hWnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
	
mov eax,uMsg
	
	.if eax==WM_MOUSEWHEEL
			invoke GetProp,hWnd,addr SBAR
			test eax,eax
			jne @f
			sub esp,4
			invoke EnumChildWindows,hWnd,EnumChilds,esp
			pop eax
			@@:
	
		mov ecx,wParam	
		test ecx,ecx
		js Dole
		invoke ScrollIt,hWnd,0,eax
		jmp Van
		Dole:
		invoke ScrollIt,hWnd,1,eax
		Van:	
		xor eax,eax
		ret
		
	.elseif eax==05005h
	;Broadcasted value!
	invoke SetProp,hWnd,addr ScrollLines,wParam
	ret
;	.elseif eax==WM_MOUSEMOVE
;	invoke SetFocus,hWnd
;	xor eax,eax
;	ret	
	.endif
	
	
invoke GetProp,hWnd,addr OP
invoke CallWindowProc,eax,hWnd,uMsg,wParam,lParam	
ret
NewProc endp


ScrollIt proc public uses esi hWnd:DWORD,Direction:DWORD,FromHwnd
invoke GetProp,hWnd,addr ScrollLines
mov esi,eax
test esi,esi
jne @f
or esi,1
@@:
invoke SendMessage,hWnd,WM_VSCROLL,Direction,FromHwnd
dec esi
jnz @b
ret
ScrollIt endp


EnumChilds proc public uses esi edi ebx hWnd:DWORD,UserData:DWORD

invoke TestWindow,hWnd,addr ScrollBar
test eax,eax
jne @1

	invoke GetWindowLong,hWnd,GWL_STYLE

	.if al==1 
		;RIGHT SCROLL BAR
		Invoke GetParent,hWnd
		mov esi,eax
		invoke GetProp,esi,addr SBAR
		test eax,eax
		jne Ima
		invoke SetProp,eax,addr SBAR,hWnd
		Ima:
		mov ecx,UserData
		mov eax,hWnd
		mov DWORD PTR [ecx],eax
		xor eax,eax
		ret
	.endif
		
@1:
or eax,-1	
ret
EnumChilds endp



EnumChildsFree proc public uses esi edi ebx hWnd:DWORD,UserData:DWORD

invoke TestWindow,hWnd,addr VbaWindow

test eax,eax
jne j1 ;Is VbaWindow
	
	invoke GetProp,hWnd,addr OP
	
	test UserData,-1
	jne ij2_
	;dettach from DLL
	test eax,eax
	je j1
		invoke SetWindowLong,hWnd,GWL_WNDPROC,eax
		invoke RemoveProp,hWnd,addr OP
		invoke RemoveProp,hWnd,addr ScrollLines
		jmp j1
	
	
	ij2_:
	;attach from DLL
	invoke GetWindowLong,hWnd,GWL_WNDPROC
	
	.if eax!=NewProc
		;Set New Proc!
		invoke SetWindowLong,hWnd,GWL_WNDPROC,NewProc
		invoke SetProp,hWnd,addr OP,eax
		invoke SetLineNumber,hWnd		
		sub esp,4 ;DUMMY VALUE
		invoke EnumChildWindows,hWnd,EnumChilds,esp
		add esp,4
	.endif

j1:
or eax,-1
ret
EnumChildsFree endp



EnumTop proc public uses esi edi ebx hWnd:DWORD,UserData:DWORD
sub esp,4
invoke GetWindowThreadProcessId,hWnd,esp
invoke GetCurrentProcessId
pop ecx

.if eax==ecx 
	invoke EnumChildWindows,hWnd,EnumChildsFree,UserData
.endif

or eax,-1
ret
EnumTop endp


end DllEntry
