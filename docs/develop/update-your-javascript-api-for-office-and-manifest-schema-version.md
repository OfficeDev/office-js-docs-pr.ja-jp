---
title: JavaScript API ライブラリOfficeバージョン 1.1 アドイン マニフェスト スキーマの最新バージョンへの更新
description: Office アドイン プロジェクトの JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 3405509c5bd0274b0bf36593a549dcb1389daa9f412798895958c6210869781e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079925"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>JavaScript API ライブラリOfficeバージョン 1.1 アドイン マニフェスト スキーマの最新バージョンへの更新

この記事では、Office アドイン プロジェクトに含まれる JavaScript ファイル (Office.js およびアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新する方法について説明します。

> [!NOTE]
> 2019 Visual Studioで作成されたプロジェクトでは、バージョン 1.1 が既に使用されます。 ただし、バージョン 1.1 にはマイナー アップデートがときどきあります。これは、この記事に記載されている方法を使用して適用できます。

## <a name="use-the-most-up-to-date-project-files"></a>最新のプロジェクト ファイルを使用する

Visual Studio を使用してアドインを開発する場合は、Office JavaScript API の最新の API メンバーとアドイン マニフェストの[v1.1](../develop/add-in-manifests.md)機能 (offappmanifest-1.1.xsd に対して検証される) を使用するには、Visual Studio 2019 をダウンロードする必要があります。 2019 Visual Studioダウンロードするには、「IDE のVisual Studio[を参照してください](https://visualstudio.microsoft.com/vs/)。 インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。

テキスト エディター、または Visual Studio 以外の IDE を使用してアドインを開発する場合は、Office.js に対する CDN への参照と、アドインのマニフェストで参照するスキーマのバージョンを更新する必要があります。

API およびアドイン マニフェスト機能を使用してOffice.jsを実行するには、次のコマンドを実行します。 お客様は、Office 2013 SP1 以降のオンプレミス製品を実行している必要があります。該当する場合は、SharePoint Server 2013 SP1 および関連サーバー製品、Exchange Server 2013 Service Pack 1 (SP1)、または同等のオンラインホスト製品 (Microsoft 365、SharePoint Online、および Exchange Online)。

Office、SharePoint、Exchange SP1 の各製品をダウンロードするには、次を参照してください。

- [Microsoft Office 2013 および関連のデスクトップ製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850036)

- [製品 Microsoft SharePoint Server 2013 と関連するサーバー製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850035)

- [Exchange Server 2013 Service Pack 1 の説明](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Visual Studio で作成した Office アドイン プロジェクトを更新する

Office JavaScript API およびアドイン マニフェスト スキーマの v1.1 のリリース前に作成されたプロジェクトの場合は **、NuGet パッケージ マネージャー** を使用してプロジェクトのファイルを更新し、アドインの HTML ページを更新して参照できます。 

なお、この更新プロセスは _プロジェクトごと_ に適用する必要があることに注意してください。v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返します。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>プロジェクト内Office JavaScript API ライブラリ ファイルを最新のリリースに更新する
次の手順では、ライブラリ ファイルOffice.js最新バージョンに更新します。 手順では 2019 Visual Studioを使用しますが、以前のバージョンのバージョンの場合と同様Visual Studio。

1. 2019 Visual Studioで、新しいアドイン プロジェクトを開Office **作成** します。
2. [ツール **] を**  >  **NuGet パッケージ マネージャー**  >  **ソリューションの Nuget パッケージを管理します**。
3. **[更新]** タブを選択します。
4. Microsoft.Office.js を選択します。 パッケージ ソースがパッケージ **ソースから提供** nuget.org。
5. 左側のウィンドウで、[インストール] **を選択し** 、パッケージ更新プロセスを完了します。

更新を完了するには、さらにいくつか手順を実行する必要があります。 アドインのHTML ページのヘッド タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された Office JavaScript API ライブラリを次のように参照します。

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > CDN URL の `office.js` の `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する

アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> アドイン マニフェスト スキーマのバージョンを 1.1 に更新した後、Capabilities 要素とCapability 要素を削除し、それらを Hosts [](../reference/manifest/hosts.md)要素と [](../reference/manifest/host.md) **Host** 要素、または Requirements 要素と [Requirements](specify-office-hosts-and-api-requirements.md)要素に置き換える必要があります。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>テキスト エディターまたは他の IDE で作成した Office アドイン プロジェクトを更新する

Office JavaScript API およびアドイン マニフェスト スキーマの v1.1 のリリース前に作成されたプロジェクトの場合は、アドインの HTML ページを更新して v1.1 ライブラリの CDN を参照し、スキーマ v1.1 を使用するためにアドインのマニフェスト ファイルを更新する必要があります。 

この更新プロセスは _プロジェクトごと_ に適用します。そのため、v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返す必要があります。

Office JavaScript API ファイル (Office.js およびアプリ固有の .js ファイル) のローカル コピーは必要ありません (Office.js 用の CDN を参照すると、実行時に必要なファイルがダウンロードされます)、ライブラリ ファイルのローカル コピーを使用する場合は[、NuGet Command-Line ユーティリティ](https://docs.nuget.org/consume/installing-nuget)とコマンドを使用してダウンロードできます。 `Install-Package Microsoft.Office.js`

> [!NOTE]
> v1.1 アドイン マニフェストの XSD (XML スキーマ定義) のコピーの取得については、「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)」を参照してください。


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>プロジェクト内Office JavaScript API ライブラリ ファイルを更新して、最新のリリースを使用する

1. テキスト エディターまたは IDE でアドインの HTML ページを開きます。

2. アドインのHTML ページのヘッド タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された Office JavaScript API ライブラリを次のように参照します。

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する

アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> アドイン マニフェスト スキーマのバージョンを 1.1 に更新した後、Capabilities 要素とCapability 要素を削除し、それらを Hosts [](../reference/manifest/hosts.md)要素と [](../reference/manifest/host.md) **Host** 要素、または Requirements 要素と [Requirements](specify-office-hosts-and-api-requirements.md)要素に置き換える必要があります。

## <a name="see-also"></a>関連項目

- [アプリケーションOffice API 要件を指定する](specify-office-hosts-and-api-requirements.md)]
- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- [Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
