Imports AlibreX
Imports System.Threading
Module Init
    Sub Main()
        Dim Hook As IAutomationHook
        Dim Root As IADRoot
        Try
            Dim Session As IADSession
            Dim FeatureStep As AlibreX.IADPartSession
            Hook = GetObject(, "AlibreX.AutomationHook")
            Root = Hook.Root
            Session = Root.TopmostSession
            FeatureStep = Session
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
            Hook = Nothing
            Root = Nothing
        End Try
    End Sub
End Module
