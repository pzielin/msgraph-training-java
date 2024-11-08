package graphapponlytutorial;

import com.azure.core.credential.AccessToken;
import com.azure.core.credential.TokenRequestContext;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.core.models.IProgressCallback;
import com.microsoft.graph.core.models.UploadResult;
import com.microsoft.graph.core.tasks.LargeFileUploadTask;
import com.microsoft.graph.core.tasks.PageIterator;
import com.microsoft.graph.drives.item.items.item.createuploadsession.CreateUploadSessionPostRequestBody;
import com.microsoft.graph.models.*;
import com.microsoft.graph.serviceclient.GraphServiceClient;

import java.io.File;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.StandardCharsets;
import java.util.LinkedList;
import java.util.List;
import java.util.Properties;

public class Graph {
    private static Properties _properties;
    private static ClientSecretCredential _clientSecretCredential;
    private static GraphServiceClient graphClient;

    public static void initializeGraphForAppOnlyAuth(Properties properties) throws Exception {
        // Ensure properties isn't null
        if (properties == null) {
            throw new Exception("Properties cannot be null");
        }

        _properties = properties;

        if (_clientSecretCredential == null) {
            final String clientId = _properties.getProperty("app.clientId");
            final String tenantId = _properties.getProperty("app.tenantId");
            final String clientSecret = _properties.getProperty("app.clientSecret");

            _clientSecretCredential = new ClientSecretCredentialBuilder()
                    .clientId(clientId)
                    .tenantId(tenantId)
                    .clientSecret(clientSecret)
                    .build();
        }

        if (graphClient == null) {
            graphClient = new GraphServiceClient(_clientSecretCredential,
                    new String[]{"https://graph.microsoft.com/.default"});
        }
    }


    public static String getAppOnlyToken() throws Exception {
        // Ensure credential isn't null
        if (_clientSecretCredential == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }

        // Request the .default scope as required by app-only auth
        final String[] graphScopes = new String[]{"https://graph.microsoft.com/.default"};

        final TokenRequestContext context = new TokenRequestContext();
        context.addScopes(graphScopes);

        final AccessToken token = _clientSecretCredential.getToken(context).block();
        return token.getToken();
    }

    public static UserCollectionResponse getUsers() throws Exception {
        // Ensure client isn't null
        if (graphClient == null) {
            throw new Exception("Graph has not been initialized for app-only auth");
        }

        return graphClient.users().get(requestConfig -> {
            requestConfig.queryParameters.select = new String[]{"displayName", "id", "mail"};
            requestConfig.queryParameters.top = 25;
            requestConfig.queryParameters.orderby = new String[]{"displayName"};
        });
    }

    public static void makeGraphCall() {
        findSiteId();
        listeDrivesSharePoint();
//        getImageFromSharePoint();
//        uploadSmallFileToSharePoint();
//        uploadLargeFile();
//        deleteFileToSharePoint();
//        moveFileToSharePoint();
// listeDrivesSharePoint();
    }

    public static void getImageFromSharePoint() {

        var driveId = "b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0"; // Documents
        // Extra call to get the list of items in the root repository
        DriveItemCollectionResponse result = graphClient.drives().byDriveId(driveId).items().byDriveItemId("root").children().get();
        List<DriveItem> allDriveItems = new LinkedList<>();
        PageIterator<DriveItem, DriveItemCollectionResponse> pageIterator = null;
        try {
            pageIterator = new PageIterator.Builder<DriveItem, DriveItemCollectionResponse>()
                    .client(graphClient)
                    .collectionPage(result)
                    .collectionPageFactory(DriveItemCollectionResponse::createFromDiscriminatorValue)
                    .processPageItemCallback(item -> {
                        allDriveItems.add(item);
                        return true;
                    }).build();
            pageIterator.iterate();
        } catch (Throwable th) {
            throw new RuntimeException(th);
        }
        for (DriveItem item : allDriveItems) {
            System.out.println(item.getId() + " " + item.getWebUrl());
        }
    }

    public static void uploadSmallFileToSharePoint() {
        // We'll have to decide which upload method to use depending on the size of the file
        var driveId = "b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0"; // Documents
        // Extra call to get the list of items in the root repository
        String data = "blabla";
        InputStream inputStream = new ByteArrayInputStream(data.getBytes(StandardCharsets.UTF_8));
        // id of the Ecole file 01XHULTYDCS5CK42YX5NAIQNEM6FIE44ZY
        var fileName = "test2.txt";
        var folder = "/TestFolder";
        var itemPath = "root:" + folder + "/" + fileName + ":";
        graphClient.drives().byDriveId(driveId).items().byDriveItemId(itemPath).content().put(inputStream);

        System.out.println("done");
    }

    public static void moveFileToSharePoint() {
        // We'll have to decide which upload method to use depending on the size of the file
        var driveId = "b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0"; // Documents

        DriveItem driveItem = new DriveItem();
        ItemReference parentReference = new ItemReference();
        var fileNameNew = "test2.txt";
        var folderNew = "/TestFolder2";
        parentReference.setId("01XHULTYGLWAVTGZOKX5GJFV54FAONJ7PD");
        driveItem.setParentReference(parentReference);
        driveItem.setName(fileNameNew);

        var fileName = "test2.txt";
        var folder = "/TestFolder";
        var itemPath = "root:" + folder + "/" + fileName + ":";
        graphClient.drives().byDriveId(driveId).items().byDriveItemId(itemPath).patch(driveItem);

        System.out.println("done");
    }

    public static void findSiteId() {
        graphClient.sites().get().getValue().forEach(site -> {
            System.out.println("name=" + site.getDisplayName() + " siteId"+ site.getId());
        });
    }

    public static void listeDrivesSharePoint() {
        List<Drive> drives = graphClient
                .sites()
                .bySiteId("wtv7z.sharepoint.com,c15cef79-d8cf-4938-a496-a5c3bc9da541,129d6d3a-54d4-4c67-811d-7d582d602135")
                .drives().get().getValue();
        drives.forEach(drive -> {
            System.out.println("name=" + drive.getName() +" driveId="+drive.getId());
        });

        Drive drive = graphClient
                .sites()
                .bySiteId("wtv7z.sharepoint.com,c15cef79-d8cf-4938-a496-a5c3bc9da541,129d6d3a-54d4-4c67-811d-7d582d602135")
                .drives().byDriveId("b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0").get();

        System.out.println("Drive ID: " + drive.getId());
        System.out.println("Drive Name: " + drive.getName());
        System.out.println("Drive Description: " + drive.getDescription());

        Drive drive2 = graphClient
                .drives().byDriveId("b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITV_Lr648sowR4yxa4k4c-B3").get();

        System.out.println("Drive ID: " + drive2.getId());
        System.out.println("Drive Name: " + drive2.getName());
        System.out.println("Drive Description: " + drive2.getDescription());

        System.out.println("done");
    }

    public static void deleteFileToSharePoint() {
        // We'll have to decide which upload method to use depending on the size of the file
        var driveId = "b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0"; // Documents
        // Extra call to get the list of items in the root repository
        // id of the Ecole file 01XHULTYDCS5CK42YX5NAIQNEM6FIE44ZY
        var fileName = "test2.txt";
        var folder = "/TestFolder";
        var itemPath = "root:" + folder + "/" + fileName + ":";
        graphClient.drives().byDriveId(driveId).items().byDriveItemId(itemPath).delete();
        System.out.println("done");
    }

    public static void uploadLargeFile() {
        // If the file is actually small this can fail!
        var driveId = "b!ee9cwc_YOEmklqXDvJ2lQTptnRLUVGdMgR19WC1gITWeL3hYv7cHQ7yQ1dol52R0"; // Documents
        String folderItemId = "01XHULTYEKCSLHFSMXZRAKJPOB5IRHX7G6";
        String fileName = "test.mp3";
        String filePath = "/Users/pawel/" + fileName;
        // Get an input stream for the file
        File file = new File(filePath);

        InputStream fileStream = null;
        try {
            fileStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
        long streamSize = file.length();

        // Set body of the upload session request
        CreateUploadSessionPostRequestBody uploadSessionRequest = new CreateUploadSessionPostRequestBody();
        DriveItemUploadableProperties properties = new DriveItemUploadableProperties();
        properties.getAdditionalData().put("@microsoft.graph.conflictBehavior", "replace");
        uploadSessionRequest.setItem(properties);

        // Create an upload session
        // ItemPath does not need to be a path to an existing item
        var itemPath = "root:/" + fileName + ":";
        UploadSession uploadSession = graphClient.drives()
                .byDriveId(driveId)
                .items()
                .byDriveItemId(itemPath)
                .createUploadSession()
                .post(uploadSessionRequest);

        // Create the upload task
        int maxSliceSize = 320 * 1000;
        LargeFileUploadTask<DriveItem> largeFileUploadTask = null;
        try {
            largeFileUploadTask = new LargeFileUploadTask<>(
                    graphClient.getRequestAdapter(),
                    uploadSession,
                    fileStream,
                    streamSize,
                    maxSliceSize,
                    DriveItem::createFromDiscriminatorValue);
        } catch (IllegalAccessException | IOException | InvocationTargetException | NoSuchMethodException e) {
            throw new RuntimeException(e);
        }

        int maxAttempts = 5;
        // Create a callback used by the upload provider
        IProgressCallback callback = (current, max) -> System.out.println(
                String.format("Uploaded %d bytes of %d total bytes", current, max));

        // Do the upload
        try {
            UploadResult<DriveItem> uploadResult = largeFileUploadTask.upload(maxAttempts, callback);
            if (uploadResult.isUploadSuccessful()) {
                System.out.println("Upload complete");
                System.out.println("Item ID: " + uploadResult.itemResponse.getId());
            } else {
                System.out.println("Upload failed");
            }
        } catch (IOException | InterruptedException e) {
            System.out.println("Error uploading: " + e.getMessage());
            throw new RuntimeException(e);
        }
    }
}
