/*
 * Created by SharpDevelop.
 * User: Michael
 * Date: 30/10/2015
 * Time: 11:14
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WindowFocusNotifier
{
	/// <summary>
	/// A transparent form that will track the current 'foreground' window and overlay it with a
	/// fake border, so you can easily see which window is active.  This is done as a rough-and-ready
	/// fix for the awful UI design decisions made by the MS Office team whereby it's now very difficult
	/// to tell which window has focus due to their removal of all window chrome.
	/// </summary>
	public partial class MainForm : Form
	{			
		private Win32.WinEventDelegate hook_i = null;		
		private IntPtr hLastOfficeWnd_i = IntPtr.Zero;
		private RECT lastOfficeWndRect_i = new RECT(0,0,0,0);
		private bool isMaximised_i = false;
		private bool isOffice365WindowsOnly_i = false;
		
		
		
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// Register an event hook with Windows SDK to get notified when the foreground window changes -
			// i.e. when you activate a new window.  We also have a timer that will check every few hundred
			// milliseconds, but this is more responsive.
			//
			// N.B. I tried registering a hook for EVENT_OBJECT_LOCATIONCHANGE events too, but it gave loads 
			// of firings when you just move the mouse and didn't actually fire when you move the window.
        	// I also tried EVENT_SYSTEM_MOVESIZEEND but it didn't work.
			//
			hook_i = new Win32.WinEventDelegate(WinEventProc);			
        	Win32.SetWinEventHook(Win32.EVENT_SYSTEM_FOREGROUND, Win32.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, hook_i, 0, 0, Win32.WINEVENT_OUTOFCONTEXT);        	
        	
        	//
        	// Make it display properly on startup
        	//
        	HandleWindowActivation(Win32.GetForegroundWindow());
		}
		
		
		
		
		/// <summary>
		/// Mark it as a WS_EX_TOOLWINDOW window, so the icon doesn't show up in Alt Tab
		/// </summary>
		protected override CreateParams CreateParams 
		{
			get 
			{ 
				CreateParams originalParams = base.CreateParams;
				originalParams.ExStyle |= 0x80;    // WS_EX_TOOLWINDOW  - so it doesn't show in alt-tab
				originalParams.ExStyle |= 0x80000; // WS_EX_LAYERED     - for click-through
				originalParams.ExStyle |= 0x20;    // WS_EX_TRANSPARENT - for click-through
				return originalParams;
			}
		} 
	
	    
		
	    
	    private void Debug(string sMessage_p)
	    {
	    	System.Diagnostics.Debug.WriteLine(sMessage_p);
	    }
	    
	    
	    
	    
	    private void WinEventProc(IntPtr hWinEventHook_p, uint eventType_p, IntPtr hWnd_p, int iObjectID_p, int iChild_p, uint dwEventThread_p, uint dwmsEventTime_p)
	    {
	    	HandleWindowActivation(hWnd_p);	    		    		   
	    }
	    
	    
	    
	    
	    private void HandleWindowActivation(IntPtr hWnd_p)
	    {		    	
	    	bool isBorderWindow = (hWnd_p == this.Handle);
	    	bool isTopLevelWindow = IsTopLevelWindow(hWnd_p);
	    	
	    	if (isTopLevelWindow  &&  ! isBorderWindow)
	    	{
	    		if (IsOffice365Window(hWnd_p))
			    {									
					RECT rect;
					Win32.GetWindowRect(hWnd_p, out rect);
					lastOfficeWndRect_i = rect;									
					hLastOfficeWnd_i = hWnd_p;			    	
					ResizeWindow(hWnd_p, ref rect);					
			    }
			    else
			    {
			    	hLastOfficeWnd_i = IntPtr.Zero;
			    	this.Hide();  	
			    }
	        }
	    }
	    
	
	    

	    private void ResizeWindow (IntPtr hWnd_p, ref RECT rect_p)
	    {
			WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
			Win32.GetWindowPlacement(hWnd_p, ref placement);
			
			int MARGIN_TOP = 0;
			int MARGIN_LEFT = 0;
			int MARGIN_BOTTOM = 0;
			int MARGIN_RIGHT = 0;
			
			isMaximised_i = (ShowWindowCommands.Maximized == placement.showCmd);
			
			if (isMaximised_i)
			{
				//
				//  For some reason, when maximised, the window position appears to be off the screen
				//  so need to move it in a bit.
				//
				Debug("It's maximised");
				MARGIN_TOP = 8;
				MARGIN_LEFT = 8;
				MARGIN_BOTTOM = 8;
				MARGIN_RIGHT = 8;
			}
						
			this.Location = new Point(rect_p.Left + MARGIN_LEFT, rect_p.Top + MARGIN_TOP);
			this.Width = rect_p.Width - (MARGIN_LEFT + MARGIN_RIGHT);
			this.Height = rect_p.Height - (MARGIN_BOTTOM + MARGIN_TOP);
			
			//
			//  Fix for issue #1 - set the 'top most' attribute again to make sure it doesn't go
			//  down the z-order pecking order when other top most windows appear.
			//
			this.TopMost = true;
			
		    this.Show();
	    }

	    
	    
	    private bool IsTopLevelWindow(IntPtr hWnd)
	    {
	    	return IntPtr.Zero == Win32.GetParent(hWnd);
	    }

	    
	    private string GetWindowClass(IntPtr hWnd)
	    {
	    	string sClassName = "Unknown yet...";
	    	
	    	//
		    // Pre-allocate 256 characters, since this is the maximum class name length.
		    //
		    StringBuilder className = new StringBuilder(256);
		    
			//
		    // Get the window class name
		    //
		    if (0 != Win32.GetClassName(hWnd, className, className.Capacity))
		    {
		    	sClassName = className.ToString();
		    }
		    
		    return sClassName;
	    }
	    
	    
	    	    
	    
	    public bool IsOffice365Window(IntPtr hWnd)
	    {	    		    	
	    	if (! isOffice365WindowsOnly_i)
	    	{
	    		return true;
	    	}
	    	
	    	string sClassName = GetWindowClass(hWnd);
	    	
		    string[] offendingClasses = {   "XLMAIN", // Excel
										    "VISIOA", // Visio
										    "OpusApp", // Word
										    "rctrl_renwnd32", // Outlook
										    "Framework::CFrame" }; // OneNote
		    
	    	foreach (string sOffendingClass in offendingClasses)
	    	{
	    		if (sOffendingClass.Equals(sClassName))
	    		{
	    			Debug("Window class '" + sClassName + "' is Office365");
	    			return true;
	    		}
	    	}
		    
		   	Debug("Window class '" + sClassName + "' is normal");
		    		    
		    return false;
	    }
	    
		
	    
		/// <summary>
		/// The Win32 SDK doesn't make it easy to detect when a window has moved or been resized.
		/// The simplest - if a little grubby - way around this is to have a timer repeatedly check
		/// the position of the foreground window and resize us if required.
		/// </summary>
		void Timer1Tick(object sender, EventArgs e)
		{
			IntPtr hWnd = Win32.GetForegroundWindow();
			
			if (hLastOfficeWnd_i == hWnd)
			{
				//
				// We're still on the same foreground window as before, so check to make sure
				// it hasn't moved.
				//
				RECT rect;
				Win32.GetWindowRect(hWnd, out rect);
			
				if (rect != lastOfficeWndRect_i)
				{
					Debug("Window " + hWnd + " has moved");
					lastOfficeWndRect_i = rect;
					ResizeWindow(hWnd, ref rect);
				}
			}
		}
		
		
		
		
		void MainFormPaint(object sender, PaintEventArgs e)
		{
			//
			// Cool trick from StackOverflow for making our window background totally transparent:
			// http://stackoverflow.com/questions/4314215/c-sharp-transparent-form
			//
			BackColor = Color.Lime;
			TransparencyKey = Color.Lime;
			
			//
			// Draw a coloured border around the window
			//
			Brush b = new SolidBrush(Color.Yellow);
			
			int BORDER_WIDTH = (isMaximised_i ? 3 : 5);
			int TITLE_BAR_BORDER_WIDTH = (isMaximised_i ? 23 : 27);
			
			e.Graphics.FillRectangle(b, 0, 0, BORDER_WIDTH, this.Height);  // Left vertical
			e.Graphics.FillRectangle(b, this.Width-BORDER_WIDTH, 0, BORDER_WIDTH, this.Height);  // Right vertical
			e.Graphics.FillRectangle(b, BORDER_WIDTH, 0, this.Width-BORDER_WIDTH, TITLE_BAR_BORDER_WIDTH);  // Top horizontal
			e.Graphics.FillRectangle(b, BORDER_WIDTH, this.Height - BORDER_WIDTH, this.Width-BORDER_WIDTH, BORDER_WIDTH);  // Bottom horizontal
		}
	}
}
