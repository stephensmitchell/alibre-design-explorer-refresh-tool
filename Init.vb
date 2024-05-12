Imports AlibreX
Imports System.Threading
Module Init
    Public Session As IADSession
    Public DesignSession As AlibreX.IADDesignSession
    Public FeatureStep As AlibreX.IADPartSession
    Public Hook As IAutomationHook
    Public Root As IADRoot
    Sub Main()
        Try
            Hook = GetObject(, "AlibreX.AutomationHook")
            Root = Hook.Root
            Session = Root.TopmostSession
            DesignSession = Session
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
            MsgBox(ex.Message)
        Finally
            Hook = Nothing
            Root = Nothing
        End Try
    End Sub

End Module
