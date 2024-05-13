Imports System.Runtime.InteropServices
Imports AlibreX
Imports System.Threading
Module Init
    Sub Main()
        Dim Hook As IAutomationHook = Nothing
        Dim Root As IADRoot = Nothing
        Dim Session As IADSession = Nothing
        Dim FeatureStep As IADPartSession = Nothing
        Try
            Hook = GetObject(, "AlibreX.AutomationHook")
            Root = Hook.Root
            Session = Root.TopmostSession
            FeatureStep = CType(Session, IADPartSession)
            FeatureStep.Features.Item(0).IsActive = True
            Thread.Sleep(TimeSpan.FromSeconds(0.5))
            For Each item As IADPartFeature In FeatureStep.Features
                item.IsActive = True
                Thread.Sleep(TimeSpan.FromSeconds(0.5))
            Next
            FeatureStep.RegenerateAll()
            FeatureStep.Save()
        Catch ex As ArgumentException
            Console.WriteLine(ex.Message)
        Finally
            If Not IsNothing(FeatureStep) Then
                Marshal.ReleaseComObject(FeatureStep)
                FeatureStep = Nothing
            End If
            If Not IsNothing(Session) Then
                Marshal.ReleaseComObject(Session)
                Session = Nothing
            End If
            If Not IsNothing(Root) Then
                Marshal.ReleaseComObject(Root)
                Root = Nothing
            End If
            If Not IsNothing(Hook) Then
                Marshal.ReleaseComObject(Hook)
                Hook = Nothing
            End If
        End Try
    End Sub
End Module
