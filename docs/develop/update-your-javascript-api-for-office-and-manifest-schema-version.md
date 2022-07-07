---
title: 最新の Office JavaScript API ライブラリとバージョン 1.1 アドイン マニフェスト スキーマに更新する
description: Office アドイン プロジェクトの JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32fcadb6a36ca540a799f8d6a5dfa671ee5e5de8
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660201"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>最新の Office JavaScript API ライブラリとバージョン 1.1 アドイン マニフェスト スキーマに更新する

この記事では、Office アドイン プロジェクトに含まれる JavaScript ファイル (Office.js およびアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新する方法について説明します。

> [!NOTE]
> Visual Studio 2019 で作成されたプロジェクトでは、バージョン 1.1 が既に使用されます。 ただし、バージョン 1.1 にはマイナー アップデートがときどきあります。これは、この記事に記載されている方法を使用して適用できます。

## <a name="use-the-most-up-to-date-project-files"></a>最新のプロジェクト ファイルを使用する

Visual Studio を使用してアドインを開発する場合、Office JavaScript API の最新の API メンバーとアドイン [マニフェストの v1.1 機能](../develop/add-in-manifests.md) (offappmanifest-1.1.xsd に対して検証) を使用するには、Visual Studio 2019 をダウンロードする必要があります。 Visual Studio 2019 をダウンロードするには、 [Visual Studio IDE ページ](https://visualstudio.microsoft.com/vs/)を参照してください。 インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。

Visual Studio 以外のテキスト エディターまたは IDE を使用してアドインを開発する場合は、Office.jsのコンテンツ配信ネットワーク (CDN) への参照と、アドインのマニフェストで参照されているスキーマのバージョンを更新する必要があります。

新しく更新された Office.js API とアドイン マニフェスト機能を使用して開発されたアドインを実行するには、お客様は Office 2013 SP1 以降のバージョンのオンプレミス製品を実行している必要があります。該当する場合は、SharePoint Server 2013 SP1 および関連サーバー製品、Exchange Server 2013 Service Pack 1 (SP1)、または同等のオンラインホスト製品である Microsoft 365、SharePoint Online、およびExchange Online。

Office、SharePoint、Exchange SP1 の各製品をダウンロードするには、次を参照してください。

- [Microsoft Office 2013 および関連のデスクトップ製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850036)

- [製品 Microsoft SharePoint Server 2013 と関連するサーバー製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850035)

- [Exchange Server 2013 Service Pack 1 の説明](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Visual Studio で作成した Office アドイン プロジェクトを更新する

Office JavaScript API とアドイン マニフェスト スキーマのリリース前に作成されたプロジェクトの場合は、 **NuGet パッケージ マネージャー** を使用してプロジェクトのファイルを更新し、アドインの HTML ページを更新して参照できます。

なお、この更新プロセスは _プロジェクトごと_ に適用する必要があることに注意してください。v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返します。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>プロジェクト内の Office JavaScript API ライブラリ ファイルを最新のリリースに更新する

次の手順では、Office.js ライブラリ ファイルを最新バージョンに更新します。 この手順では Visual Studio 2019 が使用されますが、以前のバージョンの Visual Studio と似ています。

1. Visual Studio 2019 で、新しい **Office アドイン** プロジェクトを開くか作成します。
2. [Tools **NuGet Package Manager** > **Manage Nuget Packages for Solution]\(ソリューションの Nuget パッケージを管理する****\** > ) を選択します。
3. **[更新]** タブを選択します。
4. Microsoft.Office.js を選択します。 パッケージ ソースが **nuget.org** からであることを確認します。
5. 左側のウィンドウで [ **インストール** ] を選択し、パッケージの更新プロセスを完了します。

更新を完了するには、さらにいくつか手順を実行する必要があります。 アドインの HTML ページの **ヘッド** タグで、既存のoffice.jsスクリプト参照をコメントアウトまたは削除し、更新された Office JavaScript API ライブラリを次のように参照します。

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > CDN URL の `office.js` の `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する

アドインのマニフェスト ファイルで、バージョン値`1.1`を変更する要素の **xmlns** 属性を **\<OfficeApp\>** 更新します (**xmlns** 属性以外の属性は変更されません)。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> アドイン マニフェスト スキーマのバージョンを 1.1 に更新した後、 **機能** と **機能** の要素を削除し、 [ホスト](/javascript/api/manifest/hosts) と [ホスト](/javascript/api/manifest/host) の要素または [要件と要件](specify-office-hosts-and-api-requirements.md)の要素に置き換える必要があります。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>テキスト エディターまたは他の IDE で作成した Office アドイン プロジェクトを更新する

Office JavaScript API とアドイン マニフェスト スキーマのリリース前に作成されたプロジェクトの場合は、アドインの HTML ページを更新して v1.1 ライブラリの CDN を参照し、アドインのマニフェスト ファイルを更新してスキーマ v1.1 を使用する必要があります。

この更新プロセスは _プロジェクトごと_ に適用します。そのため、v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返す必要があります。

Office アドインを開発するために Office JavaScript API ファイル (Office.js およびアプリ固有の.js ファイル) のローカル コピーは必要ありません (Office.js用の CDN を参照すると、実行時に必要なファイルがダウンロードされます)、ライブラリ ファイルのローカル コピーが必要な場合は [、NuGet Command-Line ユーティリティ](https://docs.nuget.org/consume/installing-nuget) とコマンドを `Install-Package Microsoft.Office.js` 使用してダウンロードできます。

> [!NOTE]
> v1.1 アドイン マニフェストの XSD (XML スキーマ定義) のコピーの取得については、「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)」を参照してください。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>最新のリリースを使用するように、プロジェクト内の Office JavaScript API ライブラリ ファイルを更新する

1. テキスト エディターまたは IDE でアドインの HTML ページを開きます。

2. アドインの HTML ページの **ヘッド** タグで、既存のoffice.jsスクリプト参照をコメントアウトまたは削除し、更新された Office JavaScript API ライブラリを次のように参照します。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する

アドインのマニフェスト ファイルで、バージョン値`1.1`を変更する要素の **xmlns** 属性を **\<OfficeApp\>** 更新します (**xmlns** 属性以外の属性は変更されません)。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> アドイン マニフェスト スキーマのバージョンを 1.1 に更新した後、 **機能** と **機能** の要素を削除し、 [ホスト](/javascript/api/manifest/hosts) と [ホスト](/javascript/api/manifest/host) の要素または [要件と要件](specify-office-hosts-and-api-requirements.md)の要素に置き換える必要があります。

## <a name="see-also"></a>関連項目

- [Office アプリケーションと API 要件を指定する](specify-office-hosts-and-api-requirements.md) ]
- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- [Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
