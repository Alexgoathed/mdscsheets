# mdscsheets


1. Abre tu hoja de cálculo en Google Sheets.
2. Ve a "Extensiones" y selecciona "Apps Script".
3. Borra cualquier código en el editor y pega el siguiente script:

```javascript
function findExactSum(target) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const expenses = sheet.getRange("A1:A" + sheet.getLastRow()).getValues().flat();
  const result = [];
  findCombinations(expenses, target, 0, [], result);

  if (result.length > 0) {
    Logger.log("Combinations that sum to " + target + ":");
    result.forEach(combo => Logger.log(combo));
  } else {
    Logger.log("No combinations found that sum to " + target);
  }
}

function findCombinations(expenses, target, startIndex, currentCombo, result) {
  if (target === 0) {
    result.push([...currentCombo]);
    return;
  }

  for (let i = startIndex; i < expenses.length; i++) {
    if (expenses[i] <= target) {
      currentCombo.push(expenses[i]);
      findCombinations(expenses, target - expenses[i], i + 1, currentCombo, result);
      currentCombo.pop();
    }
  }
}
```

4. Guarda el script con un nombre como "FindExactSum".
5. En el editor de Apps Script, haz clic en el menú desplegable junto al icono de depuración y selecciona `findExactSum`.
6. Ingresa el valor objetivo (target) que deseas buscar, por ejemplo, `findExactSum(100);` para encontrar combinaciones que sumen exactamente 100.
7. Ejecuta el script.

Este script buscará todas las combinaciones de gastos en la columna A que sumen exactamente la cantidad especificada y registrará las combinaciones en el Logger de Google Apps Script.

Sub FindExactSum()
    Dim target As Double
    Dim expenses As Variant
    Dim result As Collection

    ' Set your target sum here
    target = InputBox("Enter the target sum:", "Target Sum")

    ' Read expenses from column A
    expenses = Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row).Value

    ' Initialize result collection
    Set result = New Collection

    ' Find combinations
    Call FindCombinations(expenses, target, 1, result, Array())

    ' Display results
    If result.Count > 0 Then
        MsgBox "Combinations that sum to " & target & ":"
        For Each combo In result
            MsgBox Join(combo, ", ")
        Next combo
    Else
        MsgBox "No combinations found that sum to " & target
    End If
End Sub

Sub FindCombinations(expenses As Variant, target As Double, startIndex As Integer, result As Collection, currentCombo As Variant)
    Dim i As Integer

    If target = 0 Then
        result.Add currentCombo
        Exit Sub
    End If

    For i = startIndex To UBound(expenses)
        If expenses(i, 1) <= target Then
            Call FindCombinations(expenses, target - expenses(i, 1), i + 1, result, AppendToArray(currentCombo, expenses(i, 1)))
        End If
    Next i
End Sub

Function AppendToArray(arr As Variant, value As Variant) As Variant
    Dim newArr() As Variant
    Dim i As Integer

    If IsEmpty(arr) Then
        ReDim newArr(0)
        newArr(0) = value
    Else
        ReDim newArr(UBound(arr) + 1)
        For i = 0 To UBound(arr)
            newArr(i) = arr(i)
        Next i
        newArr(UBound(arr) + 1) = value
    End If

    AppendToArray = newArr
End Function


