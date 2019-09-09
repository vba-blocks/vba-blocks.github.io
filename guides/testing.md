# Testing

With `vba test`, vba-blocks uses the convention of calling `Tests.Run` in your project, passing a file path as the first argument. Anything written to the given file will be displayed in the command line / terminal.

1. Add a new module `Tests` to your project
2. Create a new `Public Sub Run()` in `Tests`
3. Add a `FilePath As String` argument to `Run`

```vb
' Tests.bas
Public Sub Run(FilePath As String)
  Open FilePath For Append As #1
  Print #1, "It works!"
  Close #1
End Sub
```

You can use [vba-test](https://github.com/VBA-tools/vba-test) to test your project with vba-blocks.

1. Add `vba-tools/test` to your project
2. Add a `FileReporter` to `Run`
3. Test your code with `TestSuite` and `TestCase`

```toml
# vba-block.toml
# ...

[dependencies]
"vba-tools/test" = "^2"
```

```vb
' Tests.bas
Public Sub Run(FilePath As String)
  Dim Suite As New TestSuite
  Suite.Description = "(project name)"

  Dim Reporter As New FileReporter
  Reporter.WriteTo FilePath
  Reporter.ListenTo Suite

  ' Test your code...
End Sub
```
