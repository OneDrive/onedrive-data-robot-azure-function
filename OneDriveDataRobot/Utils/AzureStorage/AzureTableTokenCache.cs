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
    using System;
    using Microsoft.IdentityModel.Clients.ActiveDirectory;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// ADAL TokenCache implementation that stores the token cache in the provided Azure CloudTable instance.
    /// This class is shared between the Azure Function and the bootstrap project.
    /// </summary>
    public class AzureTableTokenCache : TokenCache
    {
        private readonly string signInUserId;
        private readonly CloudTable tokenCacheTable;
        
        private TokenCacheEntity cachedEntity;      // data entity stored in the Azure Table

        public AzureTableTokenCache(string userId, CloudTable cacheTable)
        {
            signInUserId = userId;
            tokenCacheTable = cacheTable;

            this.AfterAccess = AfterAccessNotification;

            cachedEntity = ReadFromTableStorage();
            if (null != cachedEntity)
            {
                Deserialize(cachedEntity.CacheBits);
            }
        }

        private TokenCacheEntity ReadFromTableStorage()
        {
            TableOperation retrieve = TableOperation.Retrieve<TokenCacheEntity>(TokenCacheEntity.PartitionKeyValue, signInUserId);
            TableResult results = tokenCacheTable.Execute(retrieve);
            return (TokenCacheEntity)results.Result;
        }

        private void AfterAccessNotification(TokenCacheNotificationArgs args)
        {
            if (this.HasStateChanged)
            {
                if (cachedEntity == null)
                {
                    cachedEntity = new TokenCacheEntity();
                }
                cachedEntity.RowKey = signInUserId;
                cachedEntity.CacheBits = Serialize();
                cachedEntity.LastWrite = DateTime.Now;

                TableOperation insert = TableOperation.InsertOrReplace(cachedEntity);
                tokenCacheTable.Execute(insert);

                this.HasStateChanged = false;
            }
        }

        /// <summary>
        /// Representation of the data stored in the Azure Table
        /// </summary>
        private class TokenCacheEntity : TableEntity
        {
            public const string PartitionKeyValue = "tokenCache";
            public TokenCacheEntity()
            {
                this.PartitionKey = PartitionKeyValue;
            }

            public byte[] CacheBits { get; set; }
            public DateTime LastWrite { get; set; }
        }

    }
}