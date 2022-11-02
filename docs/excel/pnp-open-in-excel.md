---
title: Web ページから Excel を開き、Office アドインを埋め込む
description: Web ページから Excel を開き、Office アドインを埋め込みます。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 835518fb822602d6ca1af633f96d2be1e2697f44
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810345"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Web ページから Excel を開き、Office アドインを埋め込む

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Web ページの [Excel] ボタンの画像で、アドインが埋め込まれた新しい Excel ドキュメントが開き、自動で開きます。":::

顧客が Web ページから Microsoft Excel に直接データを開くことができるように、SaaS Web アプリケーションを拡張します。 一般的なシナリオは、顧客が Web アプリケーションでデータを操作することです。 次に、データを Excel ドキュメントにコピーします。 たとえば、Excel を使用して追加の分析を実行できます。 通常、お客様は、.csv ファイルなどのファイルにデータをエクスポートし、そのデータを Excel にインポートする必要があります。 また、Office アドインをドキュメントに手動で追加する必要もあります。

Excel ドキュメントを生成して開く Web ページの 1 回のボタン クリックにステップ数を減らします。 ドキュメント内に Office アドインを埋め込み、ドキュメントが開いたときに表示することもできます。 これにより、お客様は引き続きアプリケーション機能にアクセスできます。 ドキュメントが開くと、顧客が選択したデータと Office アドインを使用して作業を続行できます。

この記事では、独自の SaaS Web アプリケーションでこのシナリオを実装するためのコードと手法について説明します。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>新しい Excel ドキュメントを作成し、Office アドインを埋め込む

まず、Web ページから Excel ドキュメントを作成し、ドキュメントにアドインを埋め込む方法について説明します。 [Office OOXML 埋め込みアドインのコード サンプル](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)は、[Script Lab アドイン](https://appsource.microsoft.com/product/office/wa104380862)を新しい Office ドキュメントに埋め込む方法を示しています。 このサンプルは Office ドキュメントで動作しますが、この記事では Excel スプレッドシートに焦点を当てます。 サンプルをビルドして実行するには、次の手順に従います。

1. からコンピューター上のフォルダーにサンプル コード  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip を抽出します。
2. サンプルをビルドして実行するには、readme の「 **プロジェクトを使用するには** 」セクションの手順に従います。
3. サンプルを実行すると、次のスクリーンショットのような Web ページが表示されます。 Web ページを使用して、開いたときにScript Labを含む新しい Excel ドキュメントを作成します。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Excel ファイルを選択し、スクリプト ラボ アドインを埋め込むための埋め込みスクリプト ラボ サンプルが表示する Web ページのスクリーン ショット。":::

### <a name="how-the-sample-works"></a>サンプルのしくみ

サンプル コードでは、OOXML SDK を使用して、選択した Excel ドキュメントにScript Lab アドインを埋め込みます。 次の情報は、readme ファイルの [ [**コード** について] セクション](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) から取得します。

**Home.aspx.cs** ファイル:

- ボタン イベント ハンドラーと基本的な UI 操作を提供します。
- 標準の ASP.NET 手法を使用して、ファイルをアップロードしてダウンロードします。
- アップロードされたファイル (xlsx、docx、または pptx) のファイル名拡張子を使用して、ファイルの種類を決定します。 Open XML SDK には通常、ファイルの種類ごとに個別の API があるため、これは最初に行う必要があります。
- **OOXMLHelper** を呼び出してファイルを検証し、**AddInEmbedder** を呼び出してファイルにScript Labを埋め込み、自動的に開くよう設定します。

**AddInEmbedder.cs** ファイル:

- このサンプルでは、Script Labを埋め込むメソッドである主要なビジネス ロジックを提供します。
- ファイルの種類に基づいて OOXML ヘルパーを呼び出します。

**OOXMLHelper.cs** ファイル:

- 詳細な OOXML 操作をすべて提供します。
- Office ファイルを検証するための標準的な手法を使用します。これは、そのファイルで **Document.Open** メソッドを呼び出すだけです。 ファイルが無効な場合、メソッドは例外をスローします。
- Open XML 2.5 SDK のリンクにある Open XML 2.5 SDK Productivity Tools によって生成された主な [コードが含](/office/open-xml/open-xml-sdk)まれています。

**OOXMLHelper.cs** ファイルの **GenerateWebExtensionPart1Content** メソッドは、Microsoft AppSource のScript Labの ID への参照を設定します。

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType** 値は、Microsoft AppSource のエイリアスである "OMEX" です。
- **ストア** の値は、Script Labの Microsoft AppSource カルチャ セクションにある "en-US" です。
- **Id** 値は、Script Labの Microsoft AppSource 資産 ID です。

自動開くファイル共有カタログからアドインを設定する場合は、さまざまな値を使用します。

**StoreType** 値は "FileSystem" です。

- **[ストア]** の値はネットワーク共有の URL です。たとえば、"\\\\MyComputer\\MySharedFolder" です。 これは、Office セキュリティ センターで共有の信頼されたカタログ アドレスとして表示される正確な URL である必要があります。
- **Id** 値は、アドイン マニフェストのアプリ ID です。
> [!NOTE]
> これらの属性の代替値の詳細については、「ドキュメントを使用 [して作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)」を参照してください。

## <a name="use-the-fluent-ui"></a>Fluent UI を使用する

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel、PowerPoint の Fluent UI アイコン。":::

ベスト プラクティスは、Fluent UI を使用して、ユーザーが Microsoft 製品間を移行できるようにすることです。 常に Office アイコンを使用して、Web ページから起動する Office アプリケーションを示す必要があります。 Excel アイコンを使用して Excel アプリケーションを起動することを示すようにサンプル コードを変更してみましょう。

1. Visual Studio でサンプルを開きます。
1. **Home.aspx** ページを開きます。
1. フォームのダウンロード ボタンである次のコードを見つけます。

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. ボタン コードを次のイメージ タグに置き換えます。

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. **F5 キーを** 押します (または **デバッグ** > **デバッグの開始)。** ホーム ページが読み込まれると、アイコンが表示されます。

詳細については、Fluent UI 開発者ポータルの [「Office ブランド アイコン](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 」を参照してください。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Excel ドキュメントを Microsoft OneDrive にアップロードする

お客様が OneDrive を使用している場合は、新しいドキュメントを OneDrive にアップロードすることをお勧めします。 これにより、ドキュメントの検索と操作が容易になります。 新しいコード サンプルを作成し、Microsoft Graph SDK を使用して新しい Excel ドキュメントを OneDrive にアップロードする方法を見てみましょう。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>クイック スタートを使用して新しい Microsoft Graph Web アプリケーションを構築する

1. に [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 移動し、手順に従って、Office サービスと対話するクイック スタート コード サンプルを作成して開きます。
1. **手順 1: 言語またはプラットフォームを選択し**、**MVC ASP.NET** 選択します。 この手順の手順では、ASP.NET MVC オプションを使用しますが、手順は任意の言語またはプラットフォームに適用されるパターンに従います。
1. **手順 2: アプリ ID とシークレットを取得し**、[**アプリ ID とシークレットの取得**] を選択します。
1. Microsoft 365 アカウントにサインインします。  
1. [ **アプリ シークレットを保存してください** ] Web ページで、アプリ シークレットを後で取得して使用できるファイルの場所に保存します。
1. [ **入手] を選択し、クイック スタートに戻ります**。
1. **手順 2: 登録に成功しました。** 生成されたアプリ シークレットを入力します。
1. **手順 3: コーディングを開始し**、[**SDK ベースのコード サンプルのダウンロード**] を選択します。
1. ダウンロード zip フォルダーをローカル フォルダーに抽出します。  
1. Visual Studio 2019 で graph-tutorial.sln ファイルを開きます。
1. ソリューションをビルドして実行し、正常に動作していることを確認します。 予定表 Web ページを使用して、Microsoft 365 予定表を表示できる必要があります。

### <a name="upload-a-file-to-onedrive"></a>OneDrive にファイルをアップロードする

1. Visual Studio 2019 で **graph-tutorial.sln** ソリューションを開き、 **PrivateSettings.config** ファイルを開きます。

1. 次のコードのように見えるように、**ida:AppScopes** キーに新しいスコープ **Files.ReadWrite** を追加します。

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. **Index.cshtml ファイルを** 開きます。
1. 次の ActionLink コードを挿入して、ファイルを OneDrive にアップロードするボタンを作成します。

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. **HomeController.cs** ファイルを開きます。
1. 次のコードを挿入して、アクション リンクからの要求を処理します。

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. **GraphHelper.cs ファイルを** 開きます。
1. 次のコードを挿入して、Microsoft Graph APIを呼び出して OneDrive に新しいファイルを作成します。

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

1. **F5 キーを** 押します (または **デバッグ** > **デバッグの開始)。** Web アプリケーションが起動します。
1. [ **ここをクリックしてサインイン] を選択** し、サインインします。
1. [ **ここをクリック] を選択して、OneDrive に新しいファイルを作成します**。
1. 新しいブラウザー タブを開き、OneDrive アカウントにサインインします。 ルート フォルダーにtest.txt ファイルが表示されます。

OneDrive にファイルをアップロードする方法を学習したので、このコードを再利用して、作成した Excel ドキュメントをアップロードできます。

## <a name="additional-considerations-for-your-solution"></a>ソリューションに関するその他の考慮事項

テクノロジとアプローチの点では、すべてのユーザーのソリューションが異なります。 次の考慮事項は、ソリューションを変更してドキュメントを開き、Office アドインを埋め込む方法を計画するのに役立ちます。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Web ページから新しい Excel スプレッドシートを作成する

このサンプルでは、既存の Excel ドキュメントを変更します。 より一般的なシナリオは、Web ページから新しい Excel スプレッドシートを作成することです。 新しいスプレッドシートを作成する方法の詳細については、「ファイル名を指定して **スプレッドシート ドキュメントを作成** する」を参照してください。 この記事では、ファイルをローカルで作成する方法を示しますが、SpreadsheetDocument.Create メソッドでオーバーロードを使用してストリーム内にファイルを作成することもできます。

### <a name="read-custom-properties-when-your-add-in-starts"></a>アドインの起動時にカスタム プロパティを読み取る

このコード サンプルでは、OOXML SDK を使用して、新しい Excel ドキュメントにスニペット ID を格納します。 Script Lab Excel ドキュメントからスニペット ID を読み取り、開いたときにそのスニペット コードを表示します。 カスタム プロパティを独自のアドイン (クエリ文字列や一時的な認証トークンなど) に送信する必要がある場合があります。アドインの起動時にカスタム プロパティを読み取る方法の詳細については、「アドインの **状態と設定の永続化** 」を参照してください。

### <a name="initialize-the-excel-document-with-data"></a>データを使用して Excel ドキュメントを初期化する

通常、顧客が Web サイトから Excel ドキュメントを開くと、そのドキュメントに Web サイトのデータが含まれていることが想定されます。 ドキュメントにデータを書き込むには、いくつかの方法があります。

- **OOXML SDK を使用してデータを書き込みます**。 SDK を使用して、ドキュメントにデータを直接書き込むことができます。 この方法は、ドキュメントを開いた瞬間にデータを使用できるようにする場合に便利です。
- **カスタム クエリ プロパティを Office アドインに渡します**。 ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタム プロパティを埋め込みます。 アドインが開くと、クエリが取得され、クエリが実行され、Office JS API を使用してクエリの結果がドキュメントに挿入されます。

### <a name="working-with-the-ooxml-sdk"></a>OOXML SDK の操作

OOXML SDK は .NET に基づいています。 Web アプリケーションが .NET を使用しない場合は、OOXML を操作する別の方法を探す必要があります。

OOXML コードを Azure 関数に配置して、.NET コードを Web アプリケーションの残りの部分から分離できます。 次に、Web アプリケーションから Azure 関数を呼び出します (Excel ドキュメントを生成します)。 Azure 関数の詳細については、「[Azure Functionsの概要](/azure/azure-functions/functions-overview)」を参照してください。

### <a name="use-single-sign-on"></a>シングル サインオンを使用する

認証を簡略化するために、アドインでシングル サインオンを実装することをお勧めします。 詳細については、「[Office アドインのシングル サインオンを有効にする](../develop/sso-in-office-add-ins.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Open XML SDK 2.5 for Office へようこそ](/office/open-xml/open-xml-sdk)
- [ドキュメントで作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)
- [アドインの状態および設定を保持する](../develop/persisting-add-in-state-and-settings.md)
- [ファイル名を指定してスプレッドシート ドキュメントを作成する](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)