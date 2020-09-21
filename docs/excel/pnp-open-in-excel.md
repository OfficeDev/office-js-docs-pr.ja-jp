---
title: Web ページから Excel を開き、Office アドインを埋め込む
description: Web ページから Excel を開き、Office アドインを埋め込みます。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 49df253c714f3ad84d2523b87e7df894b9027355
ms.sourcegitcommit: ea03e4ea2e8537d5f6d52477816209f6c1a6579c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2020
ms.locfileid: "48166930"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Web ページから Excel を開き、Office アドインを埋め込む

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Web ページ上の [Excel] ボタンのイメージアドインが埋め込まれた新しい Excel ドキュメントを開き、自動的に開きます。":::

SaaS web アプリケーションを拡張して、顧客が web ページから Microsoft Excel に直接データを開くことができるようにします。 一般的なシナリオは、ユーザーが web アプリケーション内のデータを操作することです。 その後、データを Excel ドキュメントにコピーします。 たとえば、Excel を使用して追加の分析を実行したい場合があります。 通常、お客様はデータをファイル (.csv ファイルなど) にエクスポートしてから、データを Excel にインポートする必要があります。 また、Office アドインをドキュメントに手動で追加する必要があります。

Excel ドキュメントを生成して開く、web ページ上の1回のボタンクリックに対して実行する手順の数を減らします。 また、ドキュメントの内部に Office アドインを埋め込んで、ドキュメントを開いたときに表示することもできます。 これにより、お客様は引き続きアプリケーション機能にアクセスできるようになります。 ドキュメントが開いたときに、お客様が選択したデータが、Office アドインを引き続き使用できるようになります。

この記事では、独自の SaaS web アプリケーションでこのシナリオを実装するためのコードと手法について説明します。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>新しい Excel ドキュメントを作成し、Office アドインを埋め込む

最初に、web ページから Excel ドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。 [Office OOXML Embed アドインのコードサンプル](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)は、[スクリプトラボアドイン](https://appsource.microsoft.com/product/office/wa104380862)を新しい Office ドキュメントに埋め込む方法を示しています。 このサンプルは任意の Office ドキュメントで動作しますが、この記事の Excel スプレッドシートに重点を置いて説明します。 サンプルをビルドして実行するには、次の手順を使用します。

1. サンプルコードを  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip コンピューターのフォルダーに抽出します。
2. サンプルをビルドして実行するには、に記載されている手順に従って、readme の「プロジェクト」セクション **を使用** します。
3. サンプルを実行すると、次のスクリーンショットに似た web ページが表示されます。 Web ページを使用して、スクリプトラボが含まれる新しい Excel ドキュメントを作成します。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="埋め込みスクリプトラボサンプルが表示する web ページのスクリーンショットには、Excel ファイルを選択して、スクリプトラボアドインを埋め込むことができます。":::

### <a name="how-the-sample-works"></a>サンプルの動作方法

サンプルコードでは、OOXML SDK を使用して、選択した Excel ドキュメントにスクリプトラボアドインを埋め込みます。 次の情報は、readme ファイルの [ [**コードについて** ] セクション](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) から取得されています。

ファイル **Home.aspx.cs**:

- ボタンイベントハンドラーと基本的な UI 操作を提供します。
- は、標準の ASP.NET 技法を使用してファイルをアップロードおよびダウンロードします。
- アップロードされたファイルのファイル名拡張子 (.xlsx、.docx、または .pptx) を使用して、ファイルの種類を特定します。 通常、Open XML SDK には、ファイルの種類ごとに個別の Api が含まれているため、これを最初に実行する必要があります。
- **Ooxmlhelper**を呼び出してファイルを検証し、 **AddInEmbedder**を呼び出してスクリプトラボをファイルに埋め込み、自動的に開くように設定します。

ファイル **AddInEmbedder.cs**:

- 主要なビジネスロジックを提供します。このサンプルでは、スクリプトラボを埋め込むメソッドを示します。
- ファイルの種類に基づいて、OOXML ヘルパーに呼び出しを行います。

ファイル **OOXMLHelper.cs**:

- 詳細な OOXML 操作をすべて提供します。
- Office ファイルを検証するための標準的な手法を使用します。この方法では、単に **ドキュメントの Open** メソッドを呼び出すことができます。 ファイルが無効な場合、メソッドは例外をスローします。
- Open xml 2.5 SDK 生産性ツールで生成された、 [OPEN xml 2.5 sdk](/office/open-xml/open-xml-sdk)のリンクで利用できる主なコードが含まれています。

**OOXMLHelper.cs**ファイルの**GenerateWebExtensionPart1Content**メソッドは、Microsoft Appsource の Script Lab の ID への参照を設定します。

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **Storetype**の値は、Microsoft appsource のエイリアスである "omex" です。
- **Store**値は、スクリプトラボの Microsoft appsource culture セクションにある "en-us" です。
- **Id**値は、スクリプトラボの Microsoft appsource アセット Id です。

自動開きのファイル共有カタログからアドインを設定する場合は、別の値を使用します。

**Storetype**の値は "FileSystem" です。

- **Store**値は、ネットワーク共有の URL です。たとえば、「 \\ \\ MyComputer \\ mysharedfolder」とします。 これは、Office セキュリティセンターで、共有の信頼できるカタログアドレスとして表示される正確な URL である必要があります。
- **Id**値は、アドインのマニフェストのアプリ id です。
> [!NOTE]
> これらの属性の代替値の詳細については、「文書を使用して [作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)」を参照してください。

## <a name="use-the-fluent-ui"></a>Fluent UI を使用する

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel、および PowerPoint の Fluent UI アイコン。":::

ベストプラクティスとして、Fluent UI を使用して、ユーザーが Microsoft 製品間を移行できるようにします。 Web ページから起動する Office アプリケーションを指定するには、常に Office アイコンを使用する必要があります。 Excel のアイコンを使用して Excel アプリケーションを起動することを示すように、サンプルコードを変更してみましょう。

1. Visual Studio でサンプルを開きます。
1. [ **Default.aspx** ] ページを開きます。
1. フォーム上の [ダウンロード] ボタンである次のコードを検索します。
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. ボタンコードを次のイメージタグに置き換えます。
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. **F5**キーを押します (または**デバッグ > デバッグを開始**します)。 ホームページが読み込まれると、アイコンが表示されます。

詳細については、「Fluent UI 開発者ポータルの [Office ブランドアイコン](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 」を参照してください。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Excel ドキュメントを Microsoft OneDrive にアップロードする

お客様が OneDrive を使用している場合は、OneDrive に新しいドキュメントをアップロードすることをお勧めします。 これにより、ドキュメントの検索と操作が容易になります。 新しいコードサンプルを作成し、Microsoft Graph SDK を使用して新しい Excel ドキュメントを OneDrive にアップロードする方法を確認しましょう。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>クイックスタートを使用して新しい Microsoft Graph web アプリケーションを作成する

1. に移動 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) し、手順に従って、Office 365 サービスと対話するクイックスタートのコードサンプルを作成して開きます。
1. [ **ステップ 1: 言語またはプラットフォームを選択**してください] で、[ **ASP.NET MVC**] を選択します。 この手順の手順では ASP.NET MVC オプションを使用していますが、手順は任意の言語またはプラットフォームに適用されるパターンに従います。
1. [ **手順 2: アプリ id とシークレットを取得する**] で、[ **アプリ id とシークレットを取得する**] を選択します。
1. Microsoft 365 アカウントにサインインします。  
1. [ **アプリシークレットを保存** する] web ページで、アプリシークレットを、後で取得して使用できるファイルの場所に保存します。
1. [ **取得] を選択して、クイックスタートに戻って**ください。
1. **手順 2: 登録に成功しました。** 生成されたアプリシークレットを入力します。
1. **手順 3: コーディングを開始**するには、「 **SDK ベースのコードサンプルをダウンロードする**」を選択します。
1. ダウンロードした zip フォルダーをローカルフォルダーに展開します。  
1. Visual Studio 2019 で graph-tutorial ファイルを開きます。
1. ソリューションをビルドして実行し、正しく動作していることを確認します。 予定表 web ページを使用して、Microsoft 365 の予定表を表示できるようにする必要があります。

### <a name="upload-a-file-to-onedrive"></a>OneDrive にファイルをアップロードする

1. Visual Studio 2019 で **graph-tutorial** ソリューションを開き、 **PrivateSettings.config** ファイルを開きます。
1. **Files.ReadWrite**   **Ida: appscopes**キーに新しいスコープファイルを追加して、次のコードのようにします。
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. **人差し指**ファイルを開きます。
1. OneDrive にファイルをアップロードするボタンを作成するには、次の ActionLink コードを挿入します。
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. **HomeController.cs** ファイルを開きます。
1. アクションリンクからの要求を処理するために、次のコードを挿入します。
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. **GraphHelper.cs**ファイルを開きます。
1. 次のコードを挿入して、OneDrive に新しいファイルを作成するために Microsoft Graph API を呼び出します。
    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```
1. **F5**キーを押します (または**デバッグ > デバッグを開始**します)。 Web アプリケーションが起動します。
1. **[ここをクリックしてサインイン**] を選択し、サインインします。
1. **OneDrive で新しいファイルを作成するには、[ここをクリック**します] を選択します。
1. 新しいブラウザーのタブを開いて、OneDrive アカウントにサインインします。 ルートフォルダーに test.txt ファイルが表示されます。

これで、ファイルを OneDrive にアップロードする方法を習得しました。このコードを再利用して、作成した Excel ドキュメントをアップロードすることができます。

## <a name="additional-considerations-for-your-solution"></a>ソリューションに関するその他の考慮事項

すべてのユーザーのソリューションは、テクノロジや方法によって異なります。 次の考慮事項は、ソリューションを変更してドキュメントを開いたり、Office アドインを埋め込んだりする方法を計画するのに役立ちます。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Web ページから新しい Excel スプレッドシートを作成する

このサンプルは、既存の Excel ドキュメントを変更します。 一般的なシナリオでは、web ページから新しい Excel スプレッドシートを作成します。 新しいスプレッドシートを作成する方法については、「 **スプレッドシートドキュメントを作成** する」の「ファイル名を指定する」を参照してください。 この記事では、ファイルをローカルで作成する方法について説明しますが、SpreadsheetDocument メソッドのオーバーロードを使用して、stream でファイルを作成することもできます。

### <a name="read-custom-properties-when-your-add-in-starts"></a>アドインの起動時にカスタムプロパティを読み取る

このコードサンプルでは、OOXML SDK を使用して、新しい Excel ドキュメントにスニペット ID を格納します。 スクリプトラボは、Excel ドキュメントからスニペット ID を読み取り、開いたときにスニペットコードを表示します。 独自のアドイン (クエリ文字列、一時認証トークンなど) にカスタムプロパティを送信する必要がある場合があります。アドインを開始するときにカスタムプロパティを読み取る方法について詳しくは、「 **アドインの状態と設定を永続** 化する」を参照してください。

### <a name="initialize-the-excel-document-with-data"></a>データを使用して Excel ドキュメントを初期化する

通常、顧客が web サイトから Excel ドキュメントを開くと、ドキュメントに web サイトからのデータが含まれていると予想されます。 ドキュメントにデータを書き込むには、いくつかの方法があります。

- **OOXML SDK を使用**してデータを書き込みます。 SDK を使用して、ドキュメントに任意のデータを直接書き込むことができます。 この方法は、ドキュメントが開いているときにデータを使用できるようにする場合に便利です。
- **カスタムクエリプロパティを Office アドインに渡し**ます。 ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタムプロパティを埋め込みます。 アドインが開くと、クエリが取得され、クエリが実行され、Office JS API を使用してクエリの結果がドキュメントに挿入されます。

### <a name="working-with-the-ooxml-sdk"></a>OOXML SDK を使用する

OOXML SDK は .NET に基づいています。 Web アプリケーションが .NET に対応していない場合は、OOXML を操作するための別の方法を探す必要があります。

Javascript 版の OOXML SDK は、 [javascript 用の OPEN XML sdk](https://archive.codeplex.com/?p=openxmlsdkjs)から入手できます。

OOXML コードを Azure 関数に配置して、.NET コードを web アプリケーションの他の部分と区別することができます。 その後、Web アプリケーションから Azure 関数 (Excel ドキュメントを生成するため) を呼び出します。 Azure 関数の詳細については、「 [Azure 関数の概要](https://docs.microsoft.com/azure/azure-functions/functions-overview)」を参照してください。

### <a name="simplify-authentication"></a>認証を簡略化する

通常、お客様は web アプリケーションでの作業時に認証され、サインインします。 ベストプラクティスとして、Office アドインを使用するために再度サインインする必要がないように、ドキュメントを開くときにサインインを続けることができます。 このことを適切に処理するには、短時間の認証トークンをアドインに渡します。

1. OOXML SDK を使用して、認証トークンをドキュメント内のカスタムプロパティとして保存します。
1. アドインの開始時に、ドキュメントからトークンを読み取ります。
1. これで、アドインは顧客から追加の認証手順を必要とせずに、サービスに接続できます。

> [!WARNING]
> 認証トークンをドキュメントに埋め込むと、承認されていないユーザーがトークンを入手できるセキュリティ上のリスクが生じます。 短時間の認証トークンを使用することをお勧めします。 アドインが短時間トークンを使用している場合は、ドキュメントに保存されていない新しい認証トークンをすぐに要求する必要があります。

## <a name="see-also"></a>関連項目

- [Open XML SDK 2.5 for Office へようこそ](/office/open-xml/open-xml-sdk)
- [ドキュメントで作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)
- [アドインの状態および設定を保持する](../develop/persisting-add-in-state-and-settings.md)
- [ファイル名を指定してスプレッドシート ドキュメントを作成する](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
