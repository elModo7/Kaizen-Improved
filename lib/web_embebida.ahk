/*
 written by tank updated by Jethrow
 
 AHK_L version 1.0.90.00 ansii 32 bit or higher
  
 Date : 10/18/2011
 */
#SingleInstance, Force
 main:
{
   Gosub,init

   url = %1%
   WB.Navigate(url)
   loop
      If !WB.busy
         break

   return
}

init:
{
   ;// housekeeping routines
   ;// set the tear down procedure
   OnExit,terminate
   
   ;// Create a gui
   Gui, +LastFound +Resize +OwnDialogs
   
   ;// create an instance of Internet Explorer_Server
   ;// store the iwebbrowser2 interface pointer as *WB* & the hwnd as *ATLWinHWND*
   Gui, Add, ActiveX, w1280 h600 x0 y0 vWB hwndATLWinHWND, Shell.Explorer
   
   ;// disable annoying script errors from the page
   WB.silent := true
   
   ;// necesary to accept enter and accelorator keys
   ;http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.ole.interop.ioleinplaceactiveobject(VS.80).aspx
   IOleInPlaceActiveObject_Interface:="{00000117-0000-0000-C000-000000000046}"
   
   ;// necesary to accept enter and accelorator keys
   ;// get the in place interface pointer
   pipa := ComObjQuery(WB, IOleInPlaceActiveObject_Interface)
   
   ;// necesary to accept enter and accelorator keys
   ;// capture key messages
   OnMessage(WM_KEYDOWN:=0x0100, "WM_KEYDOWN")
   OnMessage(WM_KEYUP:=0x0101, "WM_KEYDOWN")
   
   ;//Display the GUI
   gui,show, w1280 h600 , Jisho Search
   
   ;// return and allow the program
   return
}

;// capture the gui resize event
GuiSize:
{
   ;// if there is a resize event lets resize the browser
   WinMove, % "ahk_id " . ATLWinHWND, , 0,0, A_GuiWidth, A_GuiHeight
   return
}

GuiClose:
terminate:
{
   ;// housekeeping
   ;// destroy the gui
   Gui, Destroy
   ;// release the in place interface pointer
   ObjRelease(pipa)
   ExitApp
}



WM_KEYDOWN(wParam, lParam, nMsg, hWnd)
{
   global pipa
   static keys:={9:"tab", 13:"enter", 46:"delete", 38:"up", 40:"down"}
   if keys.HasKey(wParam)
   {
      WinGetClass, ClassName, ahk_id %hWnd%
      if  (ClassName = "Internet Explorer_Server")
      {
      ;// Build MSG Structure
         VarSetCapacity(Msg, 48)
         for i,val in [hWnd, nMsg, wParam, lParam, A_EventInfo, A_GuiX, A_GuiY]
            NumPut(val, Msg, (i-1)*A_PtrSize)
      ;// Call Translate Accelerator Method
         TranslateAccelerator := NumGet(NumGet(1*pipa)+5*A_PtrSize)
         DllCall(TranslateAccelerator, "Ptr",pipa, "Ptr",&Msg)
         return, 0
      }
   }
}