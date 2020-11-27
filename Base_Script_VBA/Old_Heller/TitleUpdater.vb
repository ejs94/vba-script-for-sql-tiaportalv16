Sub TitleUpdater()
Dim title
Set title = HmiRuntime.Screens(0).ScreenItems("Screen title")
title.Text = HmiRuntime.BaseScreenName
End Sub