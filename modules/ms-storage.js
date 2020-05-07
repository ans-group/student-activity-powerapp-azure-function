require('dotenv').config();
const { BlobServiceClient } = require('@azure/storage-blob');

const AZURE_STORAGE_CONNECTION_STRING=process.env.AZURE_STORAGE_CONNECTION_STRING;

module.exports = class STORAGE { 

    // Get Token Function
    async uploadBlob(containerName, blobName, blobdata) {

        try{
            // Create BlobService Client
            const blobServiceClient = await BlobServiceClient.fromConnectionString(AZURE_STORAGE_CONNECTION_STRING);

            // Get a reference to a container
            const containerClient = await blobServiceClient.getContainerClient(containerName);

            // Get a block blob client
            const blockBlobClient = containerClient.getBlockBlobClient(blobName);

            // Blob Options
            const blobOptions = { blobHTTPHeaders: { blobContentType: 'application/json' } }

            // Upload data to the blob
            const uploadBlobResponse = await blockBlobClient.upload(blobdata, blobdata.length, blobOptions);
            
            // Return Array of JSON results
            return uploadBlobResponse.requestId

        }catch(err){
            throw err;
        };

    };

};