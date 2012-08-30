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
  [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
  [Guid(GuidList.guidVS2012UIUnbreakerPkgString)]
  //VSConstants.UICONTEXT_NoSolution.ToString()
  [ProvideAutoLoad("{ADFC4E64-0397-11D1-9F4E-00A0C911004F}")]
  //Solution exists
  [ProvideAutoLoad("f1536ef8-92ec-443c-9ed7-fdadf150da82")]
  public sealed class VS2012UIUnbreakerPackage : Package, IVsSolutionLoadEvents, IVsSolutionEvents
  {

    IntPtr m_oldWinProc = IntPtr.Zero;
    bool m_insideSetRgn = false;
    IVsUIShell m_shell = null;
    IntPtr m_topHWND = IntPtr.Zero;
    NativeMethods.WndProcDelegate m_winDelg = null;

    PropertyInfo m_glowVis;
    MethodInfo m_destGlow, m_stopGlowTimer;
    Window m_wpfWin;

    UIElement m_fakeTitleBar = null;

    IntPtr SubclassWndProc(IntPtr hWnd, NativeMethods.WM msg, IntPtr wParam, IntPtr lParam)
    {
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

    /// <summary>
    /// Default constructor of the package.
    /// Inside this method you can place any initialization code that does not require 
    /// any Visual Studio service because at this point the package object is created but 
    /// not sited yet inside Visual Studio environment. The place to do all the other 
    /// initialization is the Initialize method.
    /// </summary>
    public VS2012UIUnbreakerPackage()
    {
      Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));

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
      Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
      base.Initialize();


      EnvDTE80.DTE2 dte = GetService(typeof(SDTE)) as EnvDTE80.DTE2;
      m_shell = GetService(typeof(IVsUIShell)) as IVsUIShell;
      if (m_shell != null)
        m_shell.GetDialogOwnerHwnd(out m_topHWND);

      if (HwndSource.FromHwnd(m_topHWND) == null)
      {
        m_topHWND = IntPtr.Zero;
        dte.Events.DTEEvents.OnStartupComplete += DTEEvents_OnStartupComplete; //too soon

        m_solution = ServiceProvider.GlobalProvider.GetService(typeof(SVsSolution)) as IVsSolution2;
        if (m_solution != null)
        {
          // Register for solution events
          m_solution.AdviseSolutionEvents(this, out m_solutionEventsCookie);
        }
      }
      else
      {
        ApplyHacks();
      }

      dte.Events.DTEEvents.OnBeginShutdown += DTEEvents_OnBeginShutdown;
    }


    void DTEEvents_OnBeginShutdown()
    {
      if (m_oldWinProc.ToInt64() != 0 && m_topHWND.ToInt64() != 0)
      {
        NativeMethods.SetWindowLongPtr(m_topHWND, NativeMethods.GWLP_WNDPROC, m_oldWinProc);
        HwndSource hs = HwndSource.FromHwnd(m_topHWND);
        hs.RemoveHook(new HwndSourceHook(ManagedSubClassProc));
      }
      if (m_fakeTitleBar != null)
        m_fakeTitleBar.Visibility = Visibility.Visible;
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

    void DTEEvents_OnStartupComplete()
    {
      ApplyHacks();
    }

    private void ApplyHacks()
    {
      if (m_topHWND.ToInt64() != 0)
        return; //already done

      m_shell = GetService(typeof(IVsUIShell)) as IVsUIShell;
      m_shell.GetDialogOwnerHwnd(out m_topHWND);

      HwndSource hs = HwndSource.FromHwnd(m_topHWND);
      hs.AddHook(new HwndSourceHook(ManagedSubClassProc));

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

      var bs = new StringBuilder(4096);
      NativeMethods.GetWindowText(m_topHWND, bs, bs.Capacity);
      Debug.WriteLine(bs.ToString());
      m_winDelg = new NativeMethods.WndProcDelegate(SubclassWndProc);
      m_oldWinProc = NativeMethods.SetWindowLongPtr(m_topHWND, NativeMethods.GWLP_WNDPROC, Marshal.GetFunctionPointerForDelegate(m_winDelg));
      NativeMethods.SetWindowRgn(m_topHWND, IntPtr.Zero, false);
      KillUselessGlow();

      FindFakeTitleBar(m_wpfWin);
      FrameworkElement quickSearch = FindByName(m_fakeTitleBar as FrameworkElement, "PART_GlobalSearchTitleHost");
      FrameworkElement quickSearchWeWant = FindByName(m_wpfWin, "PART_GlobalSearchMenuHost");

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

      m_fakeTitleBar.Visibility = Visibility.Collapsed;

      StopListeningForSolutionEvents();
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
        if (o.GetType().GetTypeInfo().Name == "MainWindowTitleBar")
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
      ApplyHacks();
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
      throw new NotImplementedException();
    }

    public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel)
    {
      return VSConstants.S_OK;
    }
  }
}

