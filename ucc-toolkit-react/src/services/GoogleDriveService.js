/**
 * Service to handle Google Drive operations using V3 API
 */
export const GoogleDriveService = {
    async uploadFile(accessToken, blob, fileName) {
        const metadata = {
            name: fileName,
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        };

        // Standard Google Drive Multipart Upload (requires specific format)
        // We build the body manually to ensure it's multipart/related as expected by some strict GDrive endpoints
        const boundary = 'foo_bar_baz_ucc';
        const delimiter = "\r\n--" + boundary + "\r\n";
        const close_delim = "\r\n--" + boundary + "--";

        const reader = new FileReader();
        const blobAsBase64 = await new Promise((resolve) => {
            reader.onload = () => {
                const base64 = reader.result.split(',')[1];
                resolve(base64);
            };
            reader.readAsDataURL(blob);
        });

        const multipartRequestBody =
            delimiter +
            'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
            JSON.stringify(metadata) +
            delimiter +
            'Content-Type: ' + metadata.mimeType + '\r\n' +
            'Content-Transfer-Encoding: base64\r\n\r\n' +
            blobAsBase64 +
            close_delim;

        const response = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': `multipart/related; boundary=${boundary}`,
            },
            body: multipartRequestBody,
        });

        if (!response.ok) {
            const errorData = await response.json();
            console.error('Drive upload error:', errorData);
            throw new Error(errorData.error?.message || 'Error al subir a Google Drive');
        }

        return await response.json();
    }
};
