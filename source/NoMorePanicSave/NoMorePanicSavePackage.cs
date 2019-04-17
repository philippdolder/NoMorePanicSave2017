using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Process = System.Diagnostics.Process;

namespace NoMorePanicSave
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading =true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(NoMorePanicSavePackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class NoMorePanicSavePackage : AsyncPackage
    {
        private WindowsEventHooker.WinEventDelegate otherApplicationFocusedHandlerReference;
        private WindowsEventHooker.WinEventDelegate currentInstanceFocusedHandlerReference;
        private IntPtr otherApplicationFocusedHookHandle;
        private IntPtr currentInstanceFocusedHookHandle;
        private SolutionEvents solutionEvents;
        private bool solutionOpen = true;
        private bool hasFocus;

        /// <summary>
        /// NoMorePanicSavePackage GUID string.
        /// </summary>
        public const string PackageGuidString = "22f14c01-29ae-462b-b601-171f054ed2d1";

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override async System.Threading.Tasks.Task InitializeAsync(System.Threading.CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "Initializing.");
            await base.InitializeAsync(cancellationToken, progress);
            try
            {
                var visualStudioProcess = Process.GetCurrentProcess();
                this.otherApplicationFocusedHandlerReference = this.HandleOtherApplicationFocused;
                this.currentInstanceFocusedHandlerReference = this.HandleCurrentInstanceFocused;
                this.otherApplicationFocusedHookHandle = WindowsEventHooker.SetWinEventHook(3, 3, IntPtr.Zero, this.otherApplicationFocusedHandlerReference, 0, 0, SetWinEventHookFlags.WINEVENT_OUTOFCONTEXT | SetWinEventHookFlags.WINEVENT_SKIPOWNPROCESS);
                this.currentInstanceFocusedHookHandle = WindowsEventHooker.SetWinEventHook(3, 3, IntPtr.Zero, this.currentInstanceFocusedHandlerReference, (uint)visualStudioProcess.Id, 0, SetWinEventHookFlags.WINEVENT_OUTOFCONTEXT);

                var dte = (DTE)this.GetService(typeof(DTE));

                this.solutionEvents = dte.Events.SolutionEvents;
                this.solutionEvents.BeforeClosing += this.HandleBeforeClosingSolution;
                this.solutionEvents.Opened += this.HandleSolutionOpened;

                this.GetLogger().LogInformation(this.GetPackageName(), "Initialized.");
            }
            catch (Exception exception)
            {
                this.GetLogger().LogError(this.GetPackageName(), "Exception during initialization", exception);
            }
        }

        private string GetPackageName() => this.ToString();

        protected override void Dispose(bool disposing)
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "Disposing.");

            WindowsEventHooker.UnhookWinEvent(this.currentInstanceFocusedHookHandle);
            this.currentInstanceFocusedHandlerReference = null;

            WindowsEventHooker.UnhookWinEvent(this.otherApplicationFocusedHookHandle);
            this.otherApplicationFocusedHandlerReference = null;

            this.GetLogger().LogInformation(this.GetPackageName(), "Disposed.");

            base.Dispose(disposing);
        }

        private void HandleBeforeClosingSolution()
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "Closing solution. SolutionOpen = false");
            this.solutionOpen = false;
        }

        private void HandleSolutionOpened()
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "Opened solution. SolutionOpen = true");
            this.solutionOpen = true;
        }

        private void HandleCurrentInstanceFocused(IntPtr hWinEventHook, uint eventType,
                IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "VS focused. HasFocus = true.");
            this.hasFocus = true;
        }

        private void HandleOtherApplicationFocused(
                IntPtr hWinEventHook, uint eventType,
                IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            this.GetLogger().LogInformation(this.GetPackageName(), "Other application focused.");
            if (this.hasFocus && this.solutionOpen)
            {
                this.SaveAll();

                this.hasFocus = false;
                this.GetLogger().LogInformation(this.GetPackageName(), "Saved. HasFocus = false.");
            }
        }

        private void SaveAll()
        {
            try
            {
                var dte = (DTE)this.GetService(typeof(DTE));

                dte.ExecuteCommand("File.SaveAll");
            }
            catch (Exception exception)
            {
                this.GetLogger().LogError(this.GetPackageName(), "Exception while saving", exception);
            }
        }

        private IVsActivityLog GetLogger()
        {
            return this.GetService(typeof(SVsActivityLog)) as IVsActivityLog ?? new NullLogger();
        }
    }
}
