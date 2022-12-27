using Azure.Storage.Blobs;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace DocumentLoadSanityCheckerDownload.Services
{
    public class BlobStorageService
    {
        #region Private members
        string AccessKey { get; set; }
        string ContainerName { get; set; }
        BlobContainerClient client { get; set; }
        #endregion

        #region Public methods
        public BlobStorageService(string Key, string container)
        {
            this.AccessKey = Key;
            this.ContainerName = container;
            client = new BlobContainerClient(AccessKey, ContainerName);
        }

        //Method to upload files to the blob storage
        public Uri UploadFileToBlob(string FileName, MemoryStream fileData)
        {
            var blob = client.GetBlobClient(FileName);
            var task = blob.UploadAsync(fileData, overwrite: true);
            task.Wait();
            return blob.Uri;
        }

        public MemoryStream DownloadFileFromBlob(string FileName)
        {
            var blob = client.GetBlobClient(FileName);
            MemoryStream ms = new();
            var task = blob.DownloadToAsync(ms);
            task.Wait();
            ms.Position = 0;
            return ms;
        }
        #endregion
    }
}

