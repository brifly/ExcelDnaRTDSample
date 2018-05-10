# ExcelDnaRTDSample
Example of RTD CreateObject Issue

## Reproduction

To reproduce the issue:

1. Compile and run the addin in excel
2. Open the macros for Sheet1
3. Run TestRtdServer

The macro should fail with Run-time error '424': Object Required

If you run the macro below "TestNonRtdServer" it runs fine.
The only significant difference between the two objects is that the one that fails implements IRtdServer
