/*
 * OneDrive Data Robot - Sample Code
 * Copyright (c) Microsoft Corporation
 * All rights reserved. 
 * 
 * MIT License
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy of 
 * this software and associated documentation files (the ""Software""), to deal in 
 * the Software without restriction, including without limitation the rights to use, 
 * copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
 * Software, and to permit persons to whom the Software is furnished to do so, 
 * subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all 
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
 * HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE 
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
 
namespace OneDriveDataRobot.Utils
{
    using Microsoft.WindowsAzure.Storage.Table;
    using System;
    using System.Linq;

    /// <summary>
    /// Persists information about a subscription, userId, and deltaToken state. This class is shared between the Azure Function and the bootstrap project
    /// </summary>
    public class StoredSubscriptionState : TableEntity
    {
        public StoredSubscriptionState()
        {
            this.PartitionKey = "AAA";
        }

        public string SignInUserId { get; set; }
        public string LastDeltaToken { get; set; }
        public string SubscriptionId { get; set; }

        public static StoredSubscriptionState CreateNew(string subscriptionId)
        {
            var newState = new StoredSubscriptionState();
            newState.RowKey = subscriptionId;
            newState.SubscriptionId = subscriptionId;
            return newState;
        }

        public void Insert(CloudTable table)
        {
            TableOperation insert = TableOperation.InsertOrReplace(this);
            table.Execute(insert);
        }

        public static StoredSubscriptionState Open(string subscriptionId, CloudTable table)
        {
            try
            {
                TableOperation retrieve = TableOperation.Retrieve<StoredSubscriptionState>("AAA", subscriptionId);
                TableResult results = table.Execute(retrieve);
                return (StoredSubscriptionState)results.Result;
            } catch
            {
                return null;
            }
        }

        public static StoredSubscriptionState FindUser(string userId, CloudTable table)
        {
            try
            {
                var partitionFilter = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, "AAA");
                var userIdFilter = TableQuery.GenerateFilterCondition("SignInUserId", QueryComparisons.Equal, userId);
                string filter = TableQuery.CombineFilters(partitionFilter, TableOperators.And, userIdFilter);

                var query = new TableQuery<StoredSubscriptionState>().Where(filter).Take(1);
                var matchingEntry = table.ExecuteQuery(query).FirstOrDefault();
                return matchingEntry;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error while finding existing user subscription: {ex.Message}.");
            }
            return null;
        }

        internal void Delete(CloudTable syncStateTable)
        {
            try
            {
                TableOperation remove = TableOperation.Delete(this);
                syncStateTable.Execute(remove);
            } catch { }
        }
    }

}