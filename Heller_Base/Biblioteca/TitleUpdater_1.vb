'Title updater testar se ela está funcionando.
Sub TitleUpdater_1()
Dim title
Set title = HmiRuntime.Screens(0).ScreenItems("Screen title")
title.Text = HmiRuntime.BaseScreenName
End Sub