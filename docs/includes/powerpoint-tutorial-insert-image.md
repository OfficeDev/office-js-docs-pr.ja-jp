このチュートリアルの手順では、その日の [Bing](https://www.bing.com) 写真を取得し、そのイメージをスライドに挿入します。

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="add-the-bing-photo-of-the-day-to-a-slide"></a>その日の Bing の写真をスライドに追加する

1. ソリューション エクスプローラーを使用して、**Controllers** という名前の新しいフォルダーを **HelloWorldWeb** プロジェクトに追加します。

    ![PowerPoint のチュートリアル - HelloWorldWeb プロジェクトの Controllers フォルダーを強調表示する Visual Studio ソリューション エクスプローラー ウィンドウ](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. **Controllers** フォルダーを右クリックし、**[追加] > [新規スキャフォールディング アイテム...]** を選択します。

3. **[スキャフォールディングを追加]** ダイアログ ウィンドウで、「**Web API 2 Controller - Empty**」を選択し、**[追加]** ボタンを選択します。 

4. **[コントローラーの追加]** ダイアログ ウィンドウでコントローラー名として「**PhotoController**」と入力し、**[追加]** ボタンを選択します。 Visual Studio によって **PhotoController.cs** ファイルが作成され、表示されます。

5. **PhotoController.cs** ファイルの内容全体を、Bing サービスを呼び出す次のコードに置き換え、その日の写真を Base64 でエンコードされた文字列として取得します。 Office JavaScript API を使用してイメージをドキュメントに挿入する場合は、イメージ データを Base64 でエンコードされた文字列として指定する必要があります。

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. **Home.html** ファイルで `TODO1` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[イメージの挿入]** ボタンを定義します。

    ```html
    <button class="ms-Button ms-Button--primary" id="insert-image">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Insert Image</span>
        <span class="ms-Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. **Home.js** ファイルで `TODO1` を次のコードに置き換え、**[イメージの挿入]** ボタンのイベント ハンドラーを割り当てます。

    ```js
    $('#insert-image').click(insertImage);
    ```

8. **Home.js** ファイルで `TODO2` を次のコードに置き換え、**insertImage** 関数を定義します。 この関数は Bing Web サービスからイメージをフェッチし、`insertImageFromBase64String` 関数を呼び出してそのイメージをドキュメントに挿入します。

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. **Home.js** ファイルで `TODO3` を次のコードに置き換え、`insertImageFromBase64String` 関数を定義します。 この関数は Office JavaScript API を使用してイメージをドキュメントに挿入します。 注意: 

    - `setSelectedDataAsyc` 要求の 2 番目のパラメーターとして指定されている `coercionType` オプションは、挿入されるデータの種類を示します。 

    - `asyncResult` オブジェクトは `setSelectedDataAsync` 要求が失敗した場合の状態やエラー情報など、その要求の結果をカプセル化します。

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>アドインのテスト

1. Visual Studio を使用して、新しく作成した PowerPoint アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 作業ウィンドウで、**[イメージの挿入]** ボタンを押してその日の Bing 写真を現在のスライドに追加します。

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-image-button.png)

4. Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。 アドインが停止すると、PowerPoint は自動的に閉じます。

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)