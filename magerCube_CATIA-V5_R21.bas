Attribute VB_Name = "mengerCube_auto"
Dim myProd As Product
Dim myProds As Products
Dim myDoc As ProductDocument
Dim cubeSize As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Sub menguersCube()

Set myProd = CATIA.ActiveDocument.Product
Set myProds = myProd.Products
Set myDoc = CATIA.ActiveDocument

'getCubes and put into an array
Dim cubeList As Object
Set cubeList = getCubes

'reduce size of reference cube
divideSize

'subdivide each item in array
subdivideCubes cubeList

'delete previous items
deleteCubes cubeList

End Sub
 
 Private Sub subdivideCubes(cubeList As Object)
 For Each thisCube In cubeList
    Dim refCube As Product
    Set refCube = thisCube.ReferenceProduct
    Dim initialPos(11)
    thisCube.Position.GetComponents initialPos
        For x = -1 To 1
        For y = -1 To 1
        For z = -1 To 1
            Dim copiedCube
            Set copiedCube = myProds.AddComponent(refCube)
            Dim finalPos(11)
            copiedCube.Position.GetComponents finalPos
            finalPos(9) = initialPos(9) + x * cubeSize
            finalPos(10) = initialPos(10) + y * cubeSize
            finalPos(11) = initialPos(11) + z * cubeSize
            copiedCube.Position.SetComponents finalPos
                If getDistance(initialPos, finalPos) <= cubeSize Then
                    deleteSingleCube (copiedCube.Name)
                End If
        Next
        Next
        Next

        DoEvents
        Sleep (250)
 Next

 End Sub
 
Private Sub deleteCubes(cubeList As Object)
For Each thisCube In cubeList
    cubeName = thisCube.Name
    Dim mySelection As Selection
    Set mySelection = myDoc.Selection
    mySelection.Clear
    Dim deleteCube As Product
    Set deleteCube = myProds.Item(cubeName)
    mySelection.Add deleteCube
    mySelection.Delete
        DoEvents
        Sleep (50)
Next
End Sub

Private Sub deleteSingleCube(cubeName As String)
    Dim mySelection As Selection
    Set mySelection = myDoc.Selection
    mySelection.Clear
    Dim deleteCube As Product
    Set deleteCube = myProds.Item(cubeName)
    mySelection.Add deleteCube
    mySelection.Delete
End Sub

Private Function onlyDigits(s As String) As String
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function

Private Function getCubes() As Object
Dim coll As Object
Set coll = CreateObject("System.Collections.ArrayList")
nCubes = myProds.Count
For i = 1 To nCubes
    coll.Add myProds.Item(i)
Next
Set getCubes = coll
End Function

Private Sub divideSize()
Dim length1 As Length
Set length1 = CATIA.Documents.Item("Cube.CATPart").PART.Parameters.Item("cubeSize")
length1.Value = length1.Value / 3
cubeSize = onlyDigits(length1.Value)
myProd.Update
End Sub

Private Function getDistance(initialPos, finalPos) As Double
x1 = initialPos(9)
x2 = finalPos(9)
y1 = initialPos(10)
y2 = finalPos(10)
z1 = initialPos(11)
z2 = finalPos(11)
getDistance = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2 + (z1 - z2) ^ 2)
End Function

