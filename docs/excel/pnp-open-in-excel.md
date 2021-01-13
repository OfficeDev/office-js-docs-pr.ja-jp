---
title: Web ページから Excel を開き、アドインOffice埋め込む
description: Web ページから Excel を開き、アドインOffice埋め込む。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: a88cc647fc1dba8ab6e6ddc0b504aab96517026a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839867"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a>Web ページから Excel を開き、アドインOffice埋め込む

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="アドインを埋め込み、自動開きにした新しい Excel ドキュメントを開く Web ページ上の Excel ボタンの画像。":::

SaaS Web アプリケーションを拡張して、顧客が Web ページから Microsoft Excel に直接データを開くことができる。 一般的なシナリオは、顧客が Web アプリケーションのデータを操作することです。 次に、データを Excel ドキュメントにコピーします。 たとえば、Excel を使用して追加の分析を実行できます。 通常、顧客はデータを .csv ファイルなどのファイルにエクスポートし、そのデータを Excel にインポートする必要があります。 また、ドキュメントにアドインOffice手動で追加する必要があります。

Excel ドキュメントを生成して開く Web ページのボタンを 1 回クリックする手順の数を減らします。 ドキュメント内にアドインOffice埋め込み、ドキュメントが開くと表示できます。 これにより、顧客は引き続きアプリケーション機能にアクセスできます。 ドキュメントが開くと、顧客が選択したデータとOfficeアドインは、引き続き作業できます。

この記事では、このシナリオを独自の SaaS Web アプリケーションに実装するためのコードと手法について説明します。

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a>新しい Excel ドキュメントを作成し、新Office埋め込む

最初に、Web ページから Excel ドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。 次 [Office OOXML Embed アドイン](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) のコード サンプルは [、Script Lab](https://appsource.microsoft.com/product/office/wa104380862) アドインを新しいドキュメントに埋め込むOfficeしています。 このサンプルは、すべてのドキュメントOffice動作しますが、この記事では Excel スプレッドシートに焦点を当てるだけについて説明します。 次の手順を使用して、サンプルをビルドして実行します。

1. サンプル コードをコンピューター上  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip のフォルダーに抽出します。
2. サンプルをビルドして実行するには、readme の「プロジェクトを使用するには」セクションの手順に従います。
3. サンプルを実行すると、次のスクリーン ショットのような Web ページが表示されます。 Web ページを使用して、Script Lab を含む新しい Excel ドキュメントを作成します (開きます)。
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Excel ファイルを選択してスクリプト ラボ アドインを埋め込む目的で、埋め込みスクリプト ラボ サンプルに表示される Web ページのスクリーン ショット。":::

### <a name="how-the-sample-works"></a>サンプルのしくみ

サンプル コードでは、OOXML SDK を使用して、選択した Excel ドキュメントに Script Lab アドインを埋め込む方法を示します。 次の情報は、readme ファイルの [コード [**について** ]](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) セクションから取得されます。

次の **ファイルHome.aspx.cs。**

- ボタン イベント ハンドラーと基本的な UI 操作を提供します。
- 標準的なASP.NETを使用して、ファイルをアップロードおよびダウンロードします。
- アップロードしたファイルのファイル名の拡張子 (xlsx、docx、pptx) を使用して、ファイルの種類を特定します。 Open XML SDK には通常、ファイルの種類ごとに異なる API が含まれるため、最初にこれを行う必要があります。
- **OOXMLHelper** を呼び出してファイルを検証し **、AddInEmbedder** を呼び出して Script Lab をファイルに埋め込み、自動的に開く設定を行います。

次の **ファイルAddInEmbedder.cs。**

- 主要なビジネス ロジックを提供します。このサンプルでは、Script Lab を埋め込むメソッドです。
- ファイルの種類に基づいて OOXML ヘルパーを呼び出します。

次の **ファイルOOXMLHelper.cs。**

- すべての詳細な OOXML 操作を提供します。
- ファイルに対して **Document.Open** メソッドを呼びOfficeファイルを検証するための標準的な手法を使用します。 ファイルが無効な場合、メソッドは例外をスローします。
- 主に Open XML 2.5 SDK Productivity Tools によって生成されたコードが含まれています。このコードは [、Open XML 2.5 SDK](/office/open-xml/open-xml-sdk)のリンクから参照できます。

OOXMLHelper.cs ファイルの **GenerateWebExtensionPart1Content** メソッドは、Microsoft AppSource の Script Lab の ID への参照を設定します。 

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- StoreType **値** は、Microsoft AppSource のエイリアスである "OMEX" です。
- Store **の** 値は、Script Lab の Microsoft AppSource カルチャ セクションにある "en-US" です。
- Id **の** 値は、Script Lab の Microsoft AppSource アセット ID です。

自動開き用にファイル共有カタログからアドインをセットアップする場合は、次の値を使用します。

**StoreType の値** は "FileSystem" です。

- Store **の** 値は、ネットワーク共有の URL です。たとえば \\ \\ 、「MyComputer \\ MySharedFolder」とします。 これは、セキュリティ センターで共有の信頼済みカタログ アドレスとして表示される正確な URL Office必要があります。
- Id **値** は、アドイン マニフェストのアプリ ID です。
> [!NOTE]
> これらの属性の代替値の詳細については、「ドキュメントで作業ウィンドウを自動的に [開く」を参照してください](../develop/automatically-open-a-task-pane-with-a-document.md)。

## <a name="use-the-fluent-ui"></a>Fluent UI を使用する

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel、PowerPoint の Fluent UI アイコン。":::

ベスト プラクティスは、Fluent UI を使用して、ユーザーが Microsoft 製品間を移行する場合に役立ちます。 Web ページから起動するOfficeを示Officeアイコンを常に使用する必要があります。 サンプル コードを変更して Excel アイコンを使用し、Excel アプリケーションを起動します。

1. サンプルを次のVisual Studio。
1. **Home.aspx ページを開** きます。
1. フォームのダウンロード ボタンである次のコードを検索します。
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. ボタンのコードを次のイメージ タグに置き換えます。
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. **F5 キーを押** します (またはデバッグ **>を開始します**)。 ホーム ページが読み込まれるとアイコンが表示されます。

詳しくは、Fluent UI [Officeのブランド](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) アイコンの詳細をご覧ください。  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a>Excel ドキュメントを Microsoft OneDrive にアップロードする

顧客が OneDrive を使用している場合は、OneDrive に新しいドキュメントをアップロードすることをお勧めします。 これにより、ドキュメントの検索と作業が容易になります。 新しいコード サンプルを作成し、Microsoft Graph SDK を使用して新しい Excel ドキュメントを OneDrive にアップロードする方法を確認しましょう。

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a>クイック スタートを使用して新しい Microsoft Graph Web アプリケーションを構築する

1. 手順に [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 従って、365 サービスとやり取りするクイック スタート コード サンプルを作成Office開きます。
1. 手順 **1: 言語またはプラットフォームを選択し、MVC** ASP.NET **します**。 この手順の手順では MVC オプションASP.NET使用しますが、この手順は任意の言語またはプラットフォームに適用されるパターンに従います。
1. 手順 **2: アプリ ID とシークレットを取得** し、[アプリ ID とシークレットの **取得] を選択します**。
1. Microsoft 365 アカウントにサインインします。  
1. アプリ シークレット **Web ページで** 、アプリ シークレットを後で取得して使用できるファイルの場所に保存します。
1. Choose **Got it, take me back to the quick start**.
1. 手順 **2: 登録に成功しました。** 生成されたアプリ シークレットを入力します。
1. 手順 **3: コーディングを開始し**、[SDK ベースのコード サンプルのダウンロード **] を選択します**。
1. ダウンロード zip フォルダーをローカル フォルダーに展開します。  
1. Visual Studio 2019 で graph-tutorial.sln ファイルを開きます。
1. ソリューションをビルドして実行し、正常に動作しているのを確認します。 予定表 Web ページを使用して Microsoft 365 の予定表を表示できる必要があります。

### <a name="upload-a-file-to-onedrive"></a>OneDrive にファイルをアップロードする

1. Visual Studio 2019 で **graph-tutorial.sln** ソリューションを開き、PrivateSettings.config **します。**
1. 次のコードのように、新しいスコープ **Files.ReadWrite** を   **ida:AppScopes** キーに追加します。
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. **Index.cshtml ファイルを開** きます。
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
1. アクション リンクからの要求を処理する次のコードを挿入します。
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. ファイルを **GraphHelper.cs** します。
1. 次のコードを挿入して Microsoft Graph API を呼び出し、OneDrive に新しいファイルを作成します。
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
1. **F5 キーを押** します (またはデバッグ **>を開始します**)。 Web アプリケーションが起動します。
1. Choose **Click here to sign in,** and sign in.
1. Choose **Click here to create a new file on OneDrive**.
1. 新しいブラウザー タブを開き、OneDrive アカウントにサインインします。 ルート フォルダーにtest.txtファイルが表示されます。

OneDrive にファイルをアップロードする方法を学んだので、このコードを再利用して、作成した Excel ドキュメントをアップロードできます。

## <a name="additional-considerations-for-your-solution"></a>ソリューションに関するその他の考慮事項

テクノロジとアプローチの点では、すべてのユーザーのソリューションが異なります。 次の考慮事項は、ソリューションを変更してドキュメントを開き、アドインを埋め込むOfficeに役立ちます。

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a>Web ページから新しい Excel スプレッドシートを作成する

このサンプルでは、既存の Excel ドキュメントを変更します。 より一般的なシナリオとして、Web ページから新しい Excel スプレッドシートを作成します。 新しいスプレッドシートを作成する方法の詳細については、「ファイル名を指定してスプレッドシート ドキュメントを作成する」を参照してください。 この記事では、ファイルをローカルに作成する方法を示しますが、SpreadsheetDocument.Create メソッドのオーバーロードを使用して、ファイルをストリームで作成することもできます。

### <a name="read-custom-properties-when-your-add-in-starts"></a>アドインの起動時にカスタム プロパティを読み取る

コード サンプルでは、OOXML SDK を使用して新しい Excel ドキュメントにスニペット ID を格納します。 Script Lab は、Excel ドキュメントからスニペット ID を読み取り、そのスニペット コードを開くと表示します。 カスタム プロパティを独自のアドイン (クエリ文字列、一時認証トークンなど) に送信する必要がある場合があります。アドイン **の起動時にカスタム プロパティを読** み取る方法の詳細については、「アドインの状態と設定を保持する」を参照してください。

### <a name="initialize-the-excel-document-with-data"></a>データを使用して Excel ドキュメントを初期化する

通常、顧客が Web サイトから Excel ドキュメントを開くと、そのドキュメントには Web サイトのデータが含まれると予想されます。 ドキュメントにデータを書き込むには、いくつかの方法があります。

- **OOXML SDK を使用してデータを書き込む**。 SDK を使用して、任意のデータをドキュメントに直接書き込みできます。 この方法は、ドキュメントを開いた時点でデータを使用する場合に便利です。
- **カスタム クエリ プロパティをアドインにOffice渡します**。 ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタム プロパティを埋め込む必要があります。 アドインが開くと、クエリを取得し、クエリを実行し、Office JS API を使用してクエリの結果をドキュメントに挿入します。

### <a name="working-with-the-ooxml-sdk"></a>OOXML SDK の操作

OOXML SDK は .NET に基づいて作成されています。 Web アプリケーションが .NET ではない場合は、OOXML を使用する別の方法を探す必要があります。

[Open XML SDK for JavaScript には、OOXML SDK の JavaScript バージョンが用意されています](https://archive.codeplex.com/?p=openxmlsdkjs)。

OOXML コードを Azure 関数に配置して、.NET コードを Web アプリケーションの他の部分から分離できます。 次に、Web アプリケーションから Azure 関数を呼び出します (Excel ドキュメントを生成します)。 Azure 関数について詳しくは、「Azure 関数の概要 [」をご覧ください](/azure/azure-functions/functions-overview)。

### <a name="use-single-sign-on"></a>シングル サインオンを使用する

認証を簡略化するために、アドインにシングル サインオンを実装することをお勧めします。 詳細については、「アドインのシングル サインオンを有効にする [Office参照してください。](../develop/sso-in-office-add-ins.md)

## <a name="see-also"></a>関連項目

- [Welcome to the Open XML SDK 2.5 for Office](/office/open-xml/open-xml-sdk)
- [ドキュメントで作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)
- [アドインの状態および設定を保持する](../develop/persisting-add-in-state-and-settings.md)
- [ファイル名を指定してスプレッドシート ドキュメントを作成する](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)