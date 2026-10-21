using System.Data;

namespace Flank.SsmsExtension
{
    /// <summary>
    /// A query window's connection, lent to code that disposes the
    /// connections it opens: disposing doesn't close it. While lent, SQL
    /// errors throw (see <see cref="SsmsQueryHelper.UseBorrowedConnection"/>).
    /// </summary>
    internal sealed class BorrowedConnection : IDbConnection
    {
        private readonly IDbConnection connection;
        private readonly System.Reflection.PropertyInfo? fireInfoMessages;
        private readonly bool wasFiring;

        public BorrowedConnection(IDbConnection connection)
        {
            this.connection = connection;

            fireInfoMessages = connection.GetType().GetProperty("FireInfoMessageEventOnUserErrors");
            wasFiring = fireInfoMessages != null && (bool)fireInfoMessages.GetValue(connection);

            if (wasFiring)
                fireInfoMessages!.SetValue(connection, false);
        }

        public string ConnectionString
        {
            get => connection.ConnectionString;
            set => connection.ConnectionString = value;
        }

        public int ConnectionTimeout => connection.ConnectionTimeout;

        public string Database => connection.Database;

        public ConnectionState State => connection.State;

        public IDbTransaction BeginTransaction() => connection.BeginTransaction();

        public IDbTransaction BeginTransaction(IsolationLevel il) => connection.BeginTransaction(il);

        public void ChangeDatabase(string databaseName) => connection.ChangeDatabase(databaseName);

        public IDbCommand CreateCommand() => connection.CreateCommand();

        // The query window owns it.
        public void Open()
        {
        }

        public void Close()
        {
        }

        public void Dispose()
        {
            if (wasFiring)
                fireInfoMessages!.SetValue(connection, true);
        }
    }
}
