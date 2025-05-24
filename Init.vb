Imports System.Runtime.InteropServices
Imports AlibreX
Imports System.Threading
Module Init
    Sub Main()
        Dim hook As IAutomationHook = Nothing
        Dim root As IADRoot = Nothing
        Dim session As IADSession = Nothing
        Dim partSession As IADPartSession = Nothing
        Try
            Try
                hook = CType(GetObject(, "AlibreX.AutomationHook"), IAutomationHook)
            Catch comEx As COMException
                Throw New InvalidOperationException("Alibre Design is not running. Please start Alibre Design and retry.", comEx)
            End Try
            If hook Is Nothing Then
                Throw New InvalidOperationException("Failed to obtain AutomationHook instance.")
            End If
            root = hook.Root
            If root Is Nothing Then
                Throw New InvalidOperationException("Root object is not available. Alibre Design may not be properly initialized.")
            End If
            session = root.TopmostSession
            If session Is Nothing OrElse Not TypeOf session Is IADPartSession Then
                Throw New InvalidOperationException("The current session is not a valid part session.")
            End If
            partSession = CType(session, IADPartSession)
            If partSession.Features.Count = 0 Then
                Console.WriteLine("No features found in the current session.")
            Else
                For Each feature As IADPartFeature In partSession.Features
                    feature.IsActive = True
                    Thread.Sleep(TimeSpan.FromMilliseconds(500))
                Next
                partSession.RegenerateAll()
                partSession.Save()
                Console.WriteLine("Features activated, session regenerated, and changes saved successfully.")
            End If
        Catch invalidOpEx As InvalidOperationException
            Console.WriteLine($"Operation failed: {invalidOpEx.Message}")
        Catch comEx As COMException
            Console.WriteLine($"COM error encountered: {comEx.Message}")
        Catch argEx As ArgumentException
            Console.WriteLine($"Argument error: {argEx.Message}")
        Catch ex As Exception
            Console.WriteLine($"Unexpected error: {ex.Message}")
        Finally
            If partSession IsNot Nothing Then Marshal.ReleaseComObject(partSession)
            If session IsNot Nothing Then Marshal.ReleaseComObject(session)
            If root IsNot Nothing Then Marshal.ReleaseComObject(root)
            If hook IsNot Nothing Then Marshal.ReleaseComObject(hook)
            partSession = Nothing
            session = Nothing
            root = Nothing
            hook = Nothing
        End Try
    End Sub
End Module
