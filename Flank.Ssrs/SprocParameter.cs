namespace Flank.Ssrs
{
    public sealed class SprocParameter
    {
        public string Name { get; set; }
        public string SqlType { get; set; }
        public short MaxLength { get; set; }
        public byte Precision { get; set; }
        public byte Scale { get; set; }
        public bool IsOutput { get; set; }

        /// <summary>
        /// Whether the procedure declares a default (@p int = 5). SQL Server
        /// doesn't record T-SQL defaults in its metadata, so this comes from
        /// parsing the procedure's definition (see <see cref="ProcedureDefaults"/>).
        /// </summary>
        public bool HasDefault { get; set; }

        /// <summary>
        /// The declared default as a value to show and enter (5, abc, 2024-01-01),
        /// or null if there's none or it's NULL or an expression.
        /// </summary>
        public string? DefaultValue { get; set; }

        /// <summary>The declared default as written: 5, N'abc', NULL, GETDATE(). Null if there's none.</summary>
        public string? DefaultText { get; set; }
    }
}
