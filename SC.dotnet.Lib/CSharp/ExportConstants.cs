namespace SC.dotnet.Lib.CSharp
{
    public class ExportConstants
    {
        public static class MessageBoxInfo
        {
            public const string FILE_NOT_FOUND = "File Not Found, Export File Failed.";
            public const string IO_EXCEPTION = "Read/Write Data Exception, Export File Failed.";
            public const string FORMAT_EXCEPTION = "Convert Data error, Export File Failed";
            public const string EXPORT_EXCEPTION = "Export File Failed: {0}.";
            public const string EXPORT_EXCEL_NO_DATA = "No data in report {0}.";//{0} Report name
        }

        public static class SystemDataType
        {
            public const string INT32 = "System.Int32";
            public const string DATETIME = "System.DateTime";
            public const string DOUBLE = "System.Double";
            public const string DECIMAL = "System.Decimal";
        }
    }
}
