---
title: Web Excelファイルを開き、アドインOffice埋め込む
description: Web Excelを開き、アドインにOffice埋め込む。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349379"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Web Excelファイルを開き、アドインOffice埋め込む

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="アドインを埋Excel自動開く新しいドキュメントを開く web ページExcelボタンのイメージ。":::

SaaS Web アプリケーションを拡張して、顧客が Web ページからユーザーに直接データを開Microsoft Excel。 一般的なシナリオは、顧客が Web アプリケーションのデータを操作することです。 次に、データを別のドキュメントにExcelします。 たとえば、このツールを使用して追加の分析を実行Excel。 通常、顧客はデータをファイル (.csv ファイルなど) にエクスポートし、そのデータを Excel にインポートする必要があります。 また、ドキュメントにアドインOffice手動で追加する必要があります。

ドキュメントを生成して開く Web ページのボタンを 1 回クリックする手順の数Excelします。 また、ドキュメント内にOfficeアドインを埋め込み、ドキュメントが開くと表示できます。 これにより、顧客は引き続きアプリケーション機能にアクセスできます。 ドキュメントが開くと、顧客が選択したデータと、Officeアドインが既に使用して作業を続行できます。

この記事では、このシナリオを独自の SaaS Web アプリケーションに実装するためのコードと手法について説明します。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>新しいドキュメントをExcelし、アドインOffice埋め込む

最初に、Web ページから Excelドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。 [次Office OOXML Embed アドイン](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)のコード サンプルは、新[](https://appsource.microsoft.com/product/office/wa104380862)しいドキュメントにScript Labアドインを埋め込むOffice示しています。 このサンプルは、任意のドキュメントOffice機能しますが、この記事では、Excelスプレッドシートに焦点を当てる必要があります。 次の手順を使用して、サンプルをビルドして実行します。

1. サンプル コードをコンピューター上  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip のフォルダーに抽出します。
2. サンプルをビルドして実行するには、readme の 「プロジェクトを使用するには **」セクションの** 手順に従います。
3. サンプルを実行すると、次のスクリーンショットのような Web ページが表示されます。 Web ページを使用して、開Excelを含む新Script Labドキュメントを作成します。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="埋め込みスクリプト ラボ サンプルが表示する web ページのスクリーン ショットで、Excel ファイルを選択し、スクリプト ラボ アドインを埋め込む。":::

### <a name="how-the-sample-works"></a>サンプルの動作

サンプル コードでは、OOXML SDK を使用して、Script Labを選択したドキュメントExcel埋め込みします。 次の情報は、readme ファイルの [コード [**について** ]](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) セクションから取得されます。

**Home.aspx.cs ファイル**:

- ボタン イベント ハンドラーと基本的な UI 操作を提供します。
- 標準の ASP.NET を使用して、ファイルをアップロードおよびダウンロードします。
- アップロードされたファイルのファイル名の拡張子 (xlsx、docx、または pptx) を使用して、ファイルの種類を決定します。 Open XML SDK は通常、ファイルの種類ごとに異なる API を持つため、これは最初に行う必要があります。
- **OOXMLHelper** を呼び出してファイルを検証し **、AddInEmbedder** を呼び出してファイルにScript Labを埋め込み、自動的に開く設定を行います。

ファイル **AddInEmbedder.cs:**

- 主なビジネス ロジックを提供します。このサンプルでは、このロジックを埋め込むScript Lab。
- ファイルの種類に基づいて OOXML ヘルパーを呼び出します。

ファイル **OOXMLHelper.cs**:

- すべての詳細な OOXML 操作を提供します。
- 標準の手法を使用して、Officeファイルを検証します。これは単に **Document.Open メソッドを呼** び出す方法です。 ファイルが無効な場合、メソッドは例外をスローします。
- Open [XML 2.5 SDK](/office/open-xml/open-xml-sdk)のリンクで使用できる Open XML 2.5 SDK 生産性向上ツールによって生成されたコードが主に含まれています。

**OOXMLHelper.cs** ファイルの **GenerateWebExtensionPart1Content** メソッドは、Microsoft AppSource の id Script Labを設定します。

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- **StoreType 値** は、Microsoft AppSource のエイリアスである "OMEX" です。
- ストア **の** 値は、Microsoft AppSource カルチャ セクションにある "en-US" Script Lab。
- Id **値** は、Microsoft AppSource アセット ID の値Script Lab。

ファイル共有カタログから自動開き用にアドインを設定する場合は、次の異なる値を使用します。

**StoreType 値** は "FileSystem" です。

- Store **の** 値は、ネットワーク共有の URL です。たとえば \\ \\ 、「MyComputer \\ MySharedFolder」などです。 これは、信頼センターで共有の信頼済みカタログ アドレスとして表示される正確な URL Office必要があります。
- Id **値** は、アドイン マニフェストのアプリ ID です。
> [!NOTE]
> これらの属性の代替値の詳細については、「ドキュメントを含む作業ウィンドウを自動的に [開く」を参照してください](../develop/automatically-open-a-task-pane-with-a-document.md)。

## <a name="use-the-fluent-ui"></a>UI のFluentする

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="FluentWord、Excel、およびPowerPoint。":::

ベスト プラクティスは、ユーザーが Microsoft 製品間で移行Fluent UI を使用する方法です。 Web ページから起動するアプリケーションOfficeを示Officeアイコンを常に使用する必要があります。 サンプル コードを変更して、Excelアイコンを使用して、アプリケーションを起動Excelします。

1. サンプルを [サンプル] で開Visual Studio。
1. **[Home.aspx] ページを開** きます。
1. フォームのダウンロード ボタンである次のコードを見つける。

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. ボタン コードを次のイメージ タグに置き換えます。

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. **F5 キーを押** します (**または [デバッグ] >を開始します**)。 ホーム ページが読み込まれると、アイコンが表示されます。

詳細については、「UI 開発者[ポータルOfficeブランド](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons)アイコンFluent参照してください。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>アップロードドキュメントをExcelするMicrosoft OneDrive

顧客がドキュメントを使用している場合は、OneDriveに新しいドキュメントをアップロードOneDrive。 これにより、ドキュメントの検索と作業が容易になります。 新しいコード サンプルを作成し、Microsoft Graph SDK を使用して新しいドキュメントをExcelする方法OneDrive。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>クイック スタートを使用して新しい Microsoft Graph Web アプリケーションを構築する

1. 手順に従って、クイック スタート コード サンプルを作成して開き、サービスを操作Office [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) します。
1. 手順 **1: 言語またはプラットフォームを選択し、[MVC]** を ASP.NET **します**。 この手順の手順では、MVC ASP.NET を使用しますが、この手順は、任意の言語またはプラットフォームに適用されるパターンに従います。
1. 手順 **2: アプリ ID とシークレットを取得** し、[アプリ ID **とシークレットを取得する] を選択します**。
1. アカウントにサインインMicrosoft 365します。  
1. [アプリ **シークレット Web ページを保存してください** ] で、アプリ シークレットを後で取得して使用できるファイルの場所に保存します。
1. [Got **it] を選択し、クイック スタートに戻します**。
1. 手順 **2: 登録が成功しました!** 生成されたアプリ シークレットを入力します。
1. 手順 **3: コーディングを開始し****、[SDK ベースのコード サンプルをダウンロードする] を選択します**。
1. ダウンロード zip フォルダーをローカル フォルダーに展開します。  
1. 2019 年に graph-tutorial.sln ファイルを開Visual Studioします。
1. ソリューションをビルドして実行し、正しく動作しているのを確認します。 予定表 Web ページを使用して、予定表を表示Microsoft 365があります。

### <a name="upload-a-file-to-onedrive"></a>アップロードを作成するOneDrive

1. 2019 年 2019 年に **graph-tutorial.sln** ソリューションを開きVisual Studioファイル **をPrivateSettings.config** します。
1. 次のコードのように、新しいスコープ **Files.ReadWrite** を   **ida:AppScopes** キーに追加します。

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. **Index.cshtml ファイルを開** きます。
1. 次の ActionLink コードを挿入して、ファイルをファイルにアップロードするボタンを作成OneDrive。

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. **HomeController.cs** ファイルを開きます。
1. アクション リンクからの要求を処理するには、次のコードを挿入します。

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. **GraphHelper.cs ファイルを開** きます。
1. 次のコードを挿入して、Microsoft Graph API を呼び出して、新しいファイルを作成OneDrive。

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

1. **F5 キーを押** します (**または [デバッグ] >を開始します**)。 Web アプリケーションが起動します。
1. [ **ここをクリックしてサインイン] を選択し**、サインインします。
1. [**ここをクリックして新しいファイルを作成する] をOneDrive。**
1. 新しいブラウザー タブを開き、アカウントにサインインOneDriveします。 ルート フォルダーにtest.txtファイルが表示されます。

ファイルを OneDrive にアップロードする方法を学んだので、このコードを再利用して、作成Excelドキュメントをアップロードできます。

## <a name="additional-considerations-for-your-solution"></a>ソリューションに関するその他の考慮事項

すべてのユーザーのソリューションは、テクノロジとアプローチの点で異なります。 次の考慮事項は、ソリューションを変更してドキュメントを開き、アドインを埋め込むOfficeに役立ちます。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Web ページからExcel新しいスプレッドシートを作成する

このサンプルでは、既存のドキュメントをExcelします。 より一般的なシナリオは、Web ページから新しいExcel作成する場合です。 新しいスプレッドシートを作成する方法の詳細については、「ファイル名を指定してスプレッドシート ドキュメントを作成 **する」を** 参照してください。 この記事では、ファイルをローカルに作成する方法を示しますが、SpreadsheetDocument.Create メソッドでオーバーロードを使用して、ストリーム内にファイルを作成することもできます。

### <a name="read-custom-properties-when-your-add-in-starts"></a>アドインの起動時にカスタム プロパティを読み取る

コード サンプルでは、OOXML SDK を使用して、新Excelドキュメントにスニペット ID を格納します。 Script Labドキュメントからスニペット ID を読みExcel、開くとそのスニペット コードが表示されます。 カスタム プロパティを独自のアドイン (クエリ文字列、一時的な認証トークンなど) に送信する必要がある場合があります。アドイン **の起動時にカスタム プロパティ** を読み取る方法の詳細については、「永続化アドインの状態と設定」を参照してください。

### <a name="initialize-the-excel-document-with-data"></a>データを使用Excelドキュメントを初期化する

通常、顧客が Web サイトから Excelドキュメントを開くと、そのドキュメントに Web サイトのデータが含まれると予想されます。 ドキュメントにデータを書き込むには、いくつかの方法があります。

- **OOXML SDK を使用してデータを書き込む**。 SDK を使用すると、ドキュメントに任意のデータを直接書き込みできます。 この方法は、ドキュメントを開いた瞬間にデータを使用できる場合に便利です。
- **カスタム クエリ プロパティをアドインOffice渡します**。 ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタム プロパティを埋め込む必要があります。 アドインが開くと、クエリを取得し、クエリを実行し、Office JS API を使用してクエリの結果をドキュメントに挿入します。

### <a name="working-with-the-ooxml-sdk"></a>OOXML SDK の操作

OOXML SDK は .NET に基づいて作成されます。 Web アプリケーションが .NET を使用しない場合は、OOXML を使用する別の方法を探す必要があります。

Open XML SDK for JavaScript には、OOXML SDK の [JavaScript バージョンが用意されています](https://archive.codeplex.com/?p=openxmlsdkjs)。

OOXML コードを Azure 関数に配置して、.NET コードを他の Web アプリケーションから分離できます。 次に、Web アプリケーションから Azure 関数 (Excelドキュメントを生成する) を呼び出します。 Azure 関数の詳細については、「Azure [Functions の概要」を参照してください](/azure/azure-functions/functions-overview)。

### <a name="use-single-sign-on"></a>シングル サインオンの使用

認証を簡略化するために、アドインでシングル サインオンを実装することをお勧めします。 詳細については、「Enable [single sign-on for Office アドイン」を参照してください。](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>関連項目

- [Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [ドキュメントで作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)
- [アドインの状態および設定を保持する](../develop/persisting-add-in-state-and-settings.md)
- [ファイル名を指定してスプレッドシート ドキュメントを作成する](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)