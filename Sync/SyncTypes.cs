using System;
using System.Runtime.InteropServices;

namespace VBASinc.Sync
{
    /// <summary>
    /// Closed state machine for synchronization transaction.
    /// No implicit transitions allowed.
    /// </summary>
    public enum SyncState
    {
        /// <summary>Initial state. Validation pending.</summary>
        Init = 0,

        /// <summary>Environment validated. Ready for snapshot.</summary>
        Validated = 1,

        /// <summary>Internal and external snapshots captured in memory.</summary>
        SnapshotCreated = 2,

        /// <summary>Hashes calculated, decisions made.</summary>
        Analyzed = 3,

        /// <summary>Bilateral changes detected. Abort mandatory.</summary>
        ConflictDetected = 4,

        /// <summary>Sandbox validated. Ready for commit.</summary>
        ReadyToApply = 5,

        /// <summary>Changes applied. Post-commit verification passed.</summary>
        Committed = 6,

        /// <summary>Clean state restored after error or conflict.</summary>
        RolledBack = 7,

        /// <summary>Resources released. Terminal state.</summary>
        Disposed = 8
    }

    /// <summary>
    /// Represents a detected conflict between internal and external code.
    /// </summary>
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    public class SyncConflict
    {
        public string ModuleName { get; set; } = string.Empty;
        public string InternalHash { get; set; } = string.Empty;
        public string ExternalHash { get; set; } = string.Empty;
        public string InternalContent { get; set; } = string.Empty;
        public string ExternalContent { get; set; } = string.Empty;
        public string ReasonCode { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// Result of a synchronization operation.
    /// Contains log, conflicts, and integrity certificate.
    /// </summary>
    [ComVisible(true)]
    [Guid("B2C3D4E5-F678-90AB-CDEF-123456789ABC")]
    public class SyncResult
    {
        /// <summary>True if sync completed successfully (COMMITTED state).</summary>
        public bool Success { get; set; }

        /// <summary>Final state of the transaction.</summary>
        public SyncState FinalState { get; set; }

        /// <summary>Structured log of operations (in-memory only).</summary>
        public string Log { get; set; } = string.Empty;

        /// <summary>List of unresolved conflicts.</summary>
        public SyncConflict[] Conflicts { get; set; } = Array.Empty<SyncConflict>();

        /// <summary>
        /// Integrity certificate hash.
        /// SHA-256 of final state if verification passed, empty otherwise.
        /// </summary>
        public string IntegrityCertificate { get; set; } = string.Empty;

        /// <summary>Number of modules exported (internal → external).</summary>
        public int ExportedCount { get; set; }

        /// <summary>Number of modules imported (external → internal).</summary>
        public int ImportedCount { get; set; }

        /// <summary>Error message if failed.</summary>
        public string ErrorMessage { get; set; } = string.Empty;
    }

    /// <summary>
    /// Contract violation exception.
    /// Triggers immediate rollback.
    /// </summary>
    public class ContractViolationException : Exception
    {
        public string ContractType { get; }
        public string MethodName { get; }

        public ContractViolationException(string contractType, string methodName, string message)
            : base($"[{contractType}] {methodName}: {message}")
        {
            ContractType = contractType;
            MethodName = methodName;
        }
    }
}
