using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Interop;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;
using System.Reflection;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio;

namespace Hacks.VS2012UIUnbreaker
{
  /// <summary>
  /// This is the class that implements the package exposed by this assembly.
  ///
  /// The minimum requirement for a class to be considered a valid package for Visual Studio
  /// is to implement the IVsPackage interface and register itself with the shell.
  /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
  /// to do it: it derives from the Package class that provides the implementation of the 
  /// IVsPackage interface and uses the registration attributes defined in the framework to 
  /// register itself and its components with the shell.
  /// </summary>
  // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
  // a package.
  [PackageRegistration(UseManagedResourcesOnly = true)]
  // This attribute is used to register the information needed to show this package
  // in the Help/About dialog of Visual Studio.
  [InstalledProductRegistration("#110", "#112", "1.1", IconResourceID = 400)]
  [Guid(GuidList.guidVS2012UIUnbreakerPkgString)]
  [ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids.NoSolution)]
  [ProvideAutoLoad(Microsoft.VisualStudio.Shell.Interop.UIContextGuids.SolutionExists)]
  public sealed class VS2012UIUnbreakerPackage : Package, IVsSolutionLoadEvents, IVsSolutionEvents
  {

    public class WinFixer
    {
      IntPtr m_oldWinProc = IntPtr.Zero;
      bool m_insideSetRgn = false;
      IntPtr m_topHWND = IntPtr.Zero;
      NativeMethods.WndProcDelegate m_winDelg = null;

      PropertyInfo m_glowVis;
      MethodInfo m_destGlow, m_stopGlowTimer;
      Window m_wpfWin;

      UIElement m_fakeTitleBar = null;
      VS2012UIUnbreakerPackage m_pkg;

      public WinFixer(VS2012UIUnbreakerPackage pkg, IntPtr h)
      {
        //Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering WinFixer({0})", h.ToString()));
        m_topHWND = h;
        m_pkg = pkg;
        Done = false;
        ApplyHacks();
      }

      public IntPtr HWnd { get { return m_topHWND; } }
      public bool Done { get; private set; }

      IntPtr SubclassWndProc(IntPtr hWnd, NativeMethods.WM msg, IntPtr wParam, IntPtr lParam)
      {
        if (msg == NativeMethods.WM.CLOSE || msg == NativeMethods.WM.QUIT)
          Undo();

        if ((msg == NativeMethods.WM.WINDOWPOSCHANGED || msg == NativeMethods.WM.WINDOWPOSCHANGING) && m_insideSetRgn)
        {
          return NativeMethods.DefWindowProc(hWnd, msg, wParam, lParam);
        }

        if (msg == NativeMethods.WM.NCPAINT || msg == NativeMethods.WM.NCCREATE || msg == NativeMethods.WM.NCDESTROY || msg == NativeMethods.WM.NCCALCSIZE ||
          msg == NativeMethods.WM.NCACTIVATE || msg == NativeMethods.WM.NCHITTEST || msg == NativeMethods.WM.NCMOUSEMOVE || msg == NativeMethods.WM.NCMOUSELEAVE || msg == NativeMethods.WM.NCMOUSEHOVER ||
          msg == NativeMethods.WM.NCLBUTTONDBLCLK || msg == NativeMethods.WM.NCLBUTTONDOWN || msg == NativeMethods.WM.NCLBUTTONUP ||
          msg == NativeMethods.WM.NCRBUTTONDBLCLK || msg == NativeMethods.WM.NCRBUTTONDOWN || msg == NativeMethods.WM.NCRBUTTONUP ||
          msg == NativeMethods.WM.NCMBUTTONDBLCLK || msg == NativeMethods.WM.NCMBUTTONDOWN || msg == NativeMethods.WM.NCMBUTTONUP ||
          msg == NativeMethods.WM.NCXBUTTONDBLCLK || msg == NativeMethods.WM.NCXBUTTONDOWN || msg == NativeMethods.WM.NCXBUTTONUP ||
          msg == NativeMethods.WM.DWMCOMPOSITIONCHANGED || msg == NativeMethods.WM.DWMCOLORIZATIONCOLORCHANGED || msg == NativeMethods.WM.DWMNCRENDERINGCHANGED)
        {
          if (msg == NativeMethods.WM.DWMNCRENDERINGCHANGED && wParam.ToInt64() == 0)
          {
            KillUselessGlow();
            NativeMethods.SetWindowRgn(hWnd, IntPtr.Zero, false);
            return IntPtr.Zero;
          }

          IntPtr r = NativeMethods.DefWindowProc(hWnd, msg, wParam, lParam);
          return r;
        }
        else
        {
          IntPtr r = NativeMethods.CallWindowProc(m_oldWinProc, hWnd, msg, wParam, lParam);
          return r;
        }
      }

      public void Undo()
      {
        if (m_oldWinProc.ToInt64() != 0 && m_topHWND.ToInt64() != 0)
        {
          NativeMethods.SetWindowLongPtr(m_topHWND, NativeMethods.GWLP_WNDPROC, m_oldWinProc);
          HwndSource hs = HwndSource.FromHwnd(m_topHWND);
          hs.RemoveHook(new HwndSourceHook(ManagedSubClassProc));
        }

        m_pkg.m_otherWins.Remove(this);
        if (m_pkg.m_topHWND == this)
          m_pkg.m_topHWND = null;
        //if (m_fakeTitleBar != null)
          //m_fakeTitleBar.Visibility = Visibility.Visible;
      }


      //we need to eat these but we can't do it in the native hook since the WPF internals haven't run yet
      IntPtr ManagedSubClassProc(
        IntPtr hwnd,
        int msg2,
        IntPtr wParam,
        IntPtr lParam,
        ref bool handled)
      {

        NativeMethods.WM msg = (NativeMethods.WM)msg2;
        if (msg == NativeMethods.WM.WINDOWPOSCHANGED || msg == NativeMethods.WM.WINDOWPOSCHANGING)
        {
          handled = true;
          return NativeMethods.DefWindowProc(hwnd, msg, wParam, lParam);
        }
        else
        {

          handled = false;
          return IntPtr.Zero;
        }
      }


      public void ApplyHacks()
      {
        HwndSource hs = HwndSource.FromHwnd(m_topHWND);
        if (hs == null)
        {
          Trace.WriteLine("Bail");
          throw new ArgumentException();
        }

        m_wpfWin = (Window)hs.RootVisual;
        Type winType = m_wpfWin.GetType();
        while (winType != typeof(object))
        {
          winType = winType.GetTypeInfo().BaseType;
          if (winType.GetTypeInfo().Name == "CustomChromeWindow")
            break;
        }
        m_glowVis = winType.GetProperty("IsGlowVisible", BindingFlags.NonPublic | BindingFlags.Instance);
        m_destGlow = winType.GetMethod("DestroyGlowWindows", BindingFlags.NonPublic | BindingFlags.Instance);
        m_stopGlowTimer = winType.GetMethod("StopTimer", BindingFlags.NonPublic | BindingFlags.Instance);

        if (m_pkg.m_topHWND != null && !(bool)m_glowVis.GetValue(m_wpfWin))
        {
          //Trace.WriteLine("GlowBail");
          throw new ArgumentException();
        }

        Done = true;

        hs.AddHook(new HwndSourceHook(ManagedSubClassProc));

        var bs = new StringBuilder(4096);
        NativeMethods.GetWindowText(m_topHWND, bs, bs.Capacity);
        Debug.WriteLine(bs.ToString());
        m_winDelg = new NativeMethods.WndProcDelegate(SubclassWndProc);
        m_oldWinProc = NativeMethods.SetWindowLongPtr(m_topHWND, NativeMethods.GWLP_WNDPROC, Marshal.GetFunctionPointerForDelegate(m_winDelg));
        NativeMethods.SetWindowRgn(m_topHWND, IntPtr.Zero, false);
        KillUselessGlow();

        FindFakeTitleBar(m_wpfWin);
        if (m_fakeTitleBar == null)
          return;

        FrameworkElement quickSearch = FindByName(m_fakeTitleBar as FrameworkElement, "PART_GlobalSearchTitleHost");
        FrameworkElement quickSearchWeWant = FindByName(m_wpfWin, "PART_GlobalSearchMenuHost");
        if (quickSearch != null && quickSearchWeWant != null)
        {
          DockPanel p1 = quickSearch.Parent as DockPanel;
          DockPanel p2 = quickSearchWeWant.Parent as DockPanel;
          p1.Children.Remove(quickSearch);
          foreach (FrameworkElement f in p2.Children)
          {
            Thickness t = f.Margin;
            t.Bottom = quickSearch.Margin.Bottom;
            t.Top = quickSearch.Margin.Top;
            f.Margin = t;
          }
          //last one files so we want the right order
          p2.LastChildFill = false;
          p2.Children.Add(quickSearch);
          p2.Children.Remove(quickSearchWeWant);
          p1.Children.Add(quickSearchWeWant);//jam it here so it lives
        }

        m_fakeTitleBar.Visibility = Visibility.Collapsed;
      }

      FrameworkElement FindByName(FrameworkElement root, string name)
      {
        if (root == null || root.Name == name)
          return root;

        int childCount = VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < childCount; i++)
        {
          DependencyObject o = VisualTreeHelper.GetChild(root, i);
          FrameworkElement r = FindByName(o as FrameworkElement, name);
          if (r != null)
            return r;
        }
        return null;
      }

      void FindFakeTitleBar(DependencyObject root)
      {
        if (m_fakeTitleBar != null)
        {
          return;
        }

        int childCount = VisualTreeHelper.GetChildrenCount(root);
        for (int i = 0; i < childCount; i++)
        {
          DependencyObject o = VisualTreeHelper.GetChild(root, i);
          string name = o.GetType().GetTypeInfo().Name;
          if (name == "MainWindowTitleBar" || name == "DragUndockHeader")
          {
            m_fakeTitleBar = o as UIElement;
            return;
          }
          if (m_fakeTitleBar == null)
            FindFakeTitleBar(o);
          else
            return;
        }
      }

      private void KillUselessGlow()
      {
        if (m_glowVis != null)
          m_glowVis.SetValue(m_wpfWin, false);
        if (m_destGlow != null)
          m_destGlow.Invoke(m_wpfWin, new object[] { });
        if (m_stopGlowTimer != null)
          m_stopGlowTimer.Invoke(m_wpfWin, new object[] { });
      }

    }

    public WinFixer m_topHWND = null;
    public List<WinFixer> m_otherWins = new List<WinFixer>();

    /// <summary>
    /// Default constructor of the package.
    /// Inside this method you can place any initialization code that does not require 
    /// any Visual Studio service because at this point the package object is created but 
    /// not sited yet inside Visual Studio environment. The place to do all the other 
    /// initialization is the Initialize method.
    /// </summary>
    public VS2012UIUnbreakerPackage()
    {
      Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));

      Type suctype = typeof(StringUppercaseConverter);
      FieldInfo sucfi = suctype.GetField("_suppressUppercaseConversion", BindingFlags.NonPublic | BindingFlags.Static);
      if (sucfi != null)
        sucfi.SetValue(null, new bool?(true));
    }




    /////////////////////////////////////////////////////////////////////////////
    // Overridden Package Implementation
    #region Package Members

    /// <summary>
    /// Initialization of the package; this method is called right after the package is sited, so this is the place
    /// where you can put all the initialization code that rely on services provided by VisualStudio.
    /// </summary>
    protected override void Initialize()
    {
      Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
      base.Initialize();


      EnvDTE80.DTE2 dte = GetService(typeof(SDTE)) as EnvDTE80.DTE2;
      TryAndHookup(dte);



      if (m_topHWND == null)
      {
        dte.Events.DTEEvents.OnStartupComplete += DTEEvents_OnStartupComplete; //too soon

        m_solution = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution)) as IVsSolution2;
        if (m_solution != null)
        {
          // Register for solution events
          m_solution.AdviseSolutionEvents(this, out m_solutionEventsCookie);
        }
      }

      dte.Events.WindowEvents.WindowCreated += WindowEvents_WindowCreated;
      dte.Events.WindowEvents.WindowMoved += WindowEvents_WindowMoved;
      dte.Events.DTEEvents.OnBeginShutdown += DTEEvents_OnBeginShutdown;

    }

    private void TryAndHookup(EnvDTE80.DTE2 dte)
    {
      if (m_topHWND != null)
        return;

      try
      {
        IntPtr hwnd;
        IVsUIShell s = ServiceProvider.GlobalProvider.GetService(typeof(SVsUIShell)) as IVsUIShell;
        if (s != null)
        {
          s.GetDialogOwnerHwnd(out hwnd);
          if (hwnd != IntPtr.Zero)
          {
            m_topHWND = new WinFixer(this, hwnd);
            Trace.WriteLine("SVsUIShell worked");
          }
        }
        if (m_topHWND == null)
        {
          try
          {
            m_topHWND = new WinFixer(this, (IntPtr)dte.MainWindow.HWnd);
            Trace.WriteLine("MainWindow worked");
          }
          catch (NullReferenceException)
          {

          }
        }
      }
      catch (ArgumentException)
      {

      }

      if (m_topHWND == null)
        Trace.WriteLine("Nothing!");
    }

    
    void WindowEvents_WindowMoved(EnvDTE.Window Window, int Top, int Left, int Width, int Height)
    {
      FindFloaters();
      foreach (var f in m_otherWins)
      {
        if (!f.Done)
          f.ApplyHacks();
      }
    }

    void WindowEvents_WindowCreated(EnvDTE.Window Window)
    {
      EnvDTE80.DTE2 dte = GetService(typeof(SDTE)) as EnvDTE80.DTE2;
      TryAndHookup(dte);
      FindFloaters();
    }

    void FindFloaters()
    {
      EnvDTE80.DTE2 dte = GetService(typeof(SDTE)) as EnvDTE80.DTE2;
      string solName = dte.MainWindow.Caption.Split(' ')[0];

      IntPtr hwnd = IntPtr.Zero;
      IntPtr curWin = IntPtr.Zero;
      IntPtr curProc = (IntPtr)NativeMethods.GetCurrentProcessId();
      while ((curWin = NativeMethods.FindWindowEx(IntPtr.Zero, curWin, null, null)) != IntPtr.Zero)
      {
        IntPtr windowProcess;
        uint winThread = NativeMethods.GetWindowThreadProcessId(curWin, out windowProcess);
        if (windowProcess != curProc)
          continue;

        StringBuilder windowName = new StringBuilder(4096);
        NativeMethods.GetWindowText(curWin, windowName, windowName.Capacity);
        string wnS = windowName.ToString();
        if (wnS != solName && !wnS.StartsWith(solName + " - "))
          continue;

        StringBuilder className = new StringBuilder(4096);
        NativeMethods.GetClassName(curWin, className, className.Capacity);
        if (className.ToString().StartsWith("HwndWrapper[DefaultDomain;;"))
        {
          bool found = false;
          foreach (var f in m_otherWins)
          {
            if (f.HWnd == curWin)
            {
              found = true;
              break;
            }
          }
          if (!found)
          {
            hwnd = curWin;
            break;
          }
        }

      }
      if (hwnd == IntPtr.Zero)
        return;

      try
      {
        m_otherWins.Add(new WinFixer(this, hwnd));
      }
      catch (ArgumentException)
      {

      }
    }
  

    void DTEEvents_OnBeginShutdown()
    {
      if (m_topHWND != null)
        m_topHWND.Undo();
    }

    
    void DTEEvents_OnStartupComplete()
    {
      Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering DTEEvents_OnStartupComplete() of: {0}", this.ToString()));
      EnvDTE80.DTE2 dte = GetService(typeof(SDTE)) as EnvDTE80.DTE2;
      TryAndHookup(dte);
      FindFloaters();
    }

    #endregion



    public int OnAfterBackgroundSolutionLoadComplete()
    {
      return VSConstants.S_OK;
    }

    public int OnAfterLoadProjectBatch(bool fIsBackgroundIdleBatch)
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeBackgroundSolutionLoadBegins()
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeLoadProjectBatch(bool fIsBackgroundIdleBatch)
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeOpenSolution(string pszSolutionFilename)
    {
      DTEEvents_OnStartupComplete();
      return VSConstants.S_OK;
    }

    public int OnQueryBackgroundLoadProjectBatch(out bool pfShouldDelayLoadToNextIdle)
    {
      pfShouldDelayLoadToNextIdle = false;
      return VSConstants.S_OK;
    }

    private IVsSolution2 m_solution = null;
    private uint m_solutionEventsCookie = 0;

    protected override void Dispose(bool disposing)
    {
      base.Dispose(disposing);

      StopListeningForSolutionEvents();
    }

    private void StopListeningForSolutionEvents()
    {
      // Unadvise all events
      if (m_solution != null && m_solutionEventsCookie != 0)
        m_solution.UnadviseSolutionEvents(m_solutionEventsCookie);

      m_solution = null;
      m_solutionEventsCookie = 0;
    }


    public int OnAfterCloseSolution(object pUnkReserved)
    {
      return VSConstants.S_OK;
    }

    public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy)
    {
      return VSConstants.S_OK;
    }

    public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded)
    {
      return VSConstants.S_OK;
    }

    public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved)
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeCloseSolution(object pUnkReserved)
    {
      return VSConstants.S_OK;
    }

    public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy)
    {
      return VSConstants.S_OK;
    }

    public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel)
    {
      return VSConstants.S_OK;
    }

    public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel)
    {
      return VSConstants.S_OK;
    }

    public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
    {
      return VSConstants.S_OK;
    }
  }
}

