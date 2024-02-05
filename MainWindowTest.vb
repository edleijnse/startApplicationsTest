Imports NUnit.Framework
Imports System.IO
Imports System.Threading

<TestFixture, Apartment(ApartmentState.STA)>
Public Class MainWindowTests
    Private mainWindow As MainWindow

    <SetUp>
    Public Sub Setup()
        mainWindow = New MainWindow()
    End Sub

    <Test>
    Public Sub Test_LoadButtonContentsFromIniFile()
        ' Arrange - Setting up any necessary data for the test..
        Dim inipath As String = mainWindow.GetIniFilePath()
        ' Clean up the ini file and re-create it with known data before testing
        If File.Exists(inipath) Then File.Delete(inipath)
        Using sw As StreamWriter = File.CreateText(inipath)
            sw.WriteLine("[Buttons]")
            sw.WriteLine("Button1 = Excel.exe")
            sw.WriteLine("Button2 = Word.exe")
            sw.WriteLine("Button3 = OneNote.exe")
            sw.WriteLine("Button4 = Outlook.exe")
        End Using

        ' Act - Call the method under test.
        mainWindow.LoadButtonContentsFromIniFile()

        ' Assert - Check the result against the expected results.
        ' Assert.That(mainWindow.GetButton1Content(), Is.EqualTo('Excel.exe'))
        Dim myButtonContent1 = mainWindow.GetButton1Content().ToString()
        Assert.That(myButtonContent1,[Is].EqualTo("Excel.exe"))     
        Dim myButtonContent2 = mainWindow.GetButton2Content().ToString()
        Assert.That(myButtonContent2,[Is].EqualTo("Word.exe"))     
        Dim myButtonContent3 = mainWindow.GetButton3Content().ToString()
        Assert.That(myButtonContent3,[Is].EqualTo("OneNote.exe"))  
        Dim myButtonContent4 = mainWindow.GetButton4Content().ToString()
        Assert.That(myButtonContent4,[Is].EqualTo("Outlook.exe"))  
    End Sub

    <TearDown>
    Public Sub Teardown()
        ' Clean up any objects or data after your tests run.
        mainWindow = Nothing
    End Sub
End Class