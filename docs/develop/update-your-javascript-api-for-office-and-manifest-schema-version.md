---
title: 最新の Office JavaScript API ライブラリおよびバージョン1.1 のアドインマニフェストスキーマへの更新
description: Office アドイン プロジェクトの JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: b0536b4b55accd99e002e26c467572330ba72ae2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293129"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>最新の Office JavaScript API ライブラリおよびバージョン1.1 のアドインマニフェストスキーマへの更新

この記事では、Office アドイン プロジェクトに含まれる JavaScript ファイル (Office.js およびアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新する方法について説明します。

> [!NOTE]
> Visual Studio 2019 で作成されたプロジェクトは、既にバージョン1.1 を使用しています。 ただし、バージョン 1.1 にはマイナー アップデートがときどきあります。これは、この記事に記載されている方法を使用して適用できます。

## <a name="use-the-most-up-to-date-project-files"></a>最新のプロジェクト ファイルを使用する

Visual Studio を使用してアドインを開発する場合は、Office JavaScript API の最新の API メンバーと [アドインマニフェストの v2.0 機能](../develop/add-in-manifests.md) (offappmanifest-に対して検証されます) を使用するには、visual studio 2019 をダウンロードする必要があります。 Visual Studio 2019 をダウンロードするには、 [Visual STUDIO IDE ページ](https://visualstudio.microsoft.com/vs/)を参照してください。 インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。

テキスト エディター、または Visual Studio 以外の IDE を使用してアドインを開発する場合は、Office.js に対する CDN への参照と、アドインのマニフェストで参照するスキーマのバージョンを更新する必要があります。

新規および更新された Office.js API およびアドインマニフェスト機能を使用して開発したアドインを実行するには、Office 2013 SP1 以降のバージョンのオンプレミス製品を実行している必要があり365ます。また、該当する場合は、SharePoint Server 2013 SP1 および関連するサーバー製品、Exchange Server 2013 Service Pack 1 (SP1)、またはそれと同等の

Office、SharePoint、Exchange SP1 の各製品をダウンロードするには、次を参照してください。

- [Microsoft Office 2013 および関連のデスクトップ製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850036)

- [製品 Microsoft SharePoint Server 2013 と関連するサーバー製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧](https://support.microsoft.com/kb/2850035)

- [Exchange Server 2013 Service Pack 1 の説明](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Visual Studio で作成した Office アドイン プロジェクトを更新する

Office JavaScript API およびアドインマニフェストスキーマのリリースの前に作成されたプロジェクトでは、 **NuGet パッケージマネージャー**を使用してプロジェクトのファイルを更新し、アドインの HTML ページを更新してそれらを参照することができます。 

なお、この更新プロセスは _プロジェクトごと_ に適用する必要があることに注意してください。v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返します。

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>プロジェクト内の Office JavaScript API ライブラリファイルを最新のリリースに更新する
次の手順では、Office.js ライブラリファイルを最新バージョンに更新します。 手順では Visual Studio 2019 を使用していますが、以前のバージョンの Visual Studio に似ています。

1. Visual Studio 2019 で、 **Office アドイン** プロジェクトを開くか新規作成します。
2. **ツール**  >  の選択**nuget パッケージマネージャー**  >  **ソリューションの nuget パッケージを管理**します。
3. **[更新]** タブを選択します。
4. Microsoft.Office.js を選択します。 パッケージソースが **nuget.org**からのものであることを確認します。
5. 左側のウィンドウで、[ **インストール** ] を選択し、パッケージの更新プロセスを完了します。

更新を完了するには、さらにいくつか手順を実行する必要があります。 アドインの HTML ページの **head** タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された OFFICE JavaScript API ライブラリを次のように参照します。

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
> アドインマニフェストスキーマのバージョンを1.1 に更新した後で、 **機能** 要素と **機能** 要素を削除し、それらを [Hosts](../reference/manifest/hosts.md) 要素と [Host](../reference/manifest/host.md) 要素、または [要件と要件要素](specify-office-hosts-and-api-requirements.md)のいずれかに置き換える必要があります。

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>テキスト エディターまたは他の IDE で作成した Office アドイン プロジェクトを更新する

Office JavaScript API およびアドインマニフェストスキーマのリリースの前に作成されたプロジェクトでは、アドインの HTML ページを更新して v2.0 ライブラリの CDN を参照し、アドインのマニフェストファイルを更新してスキーマ v1.1 を使用する必要があります。 

この更新プロセスは_プロジェクトごと_に適用します。そのため、v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返す必要があります。

Office の JavaScript API ファイル (Office.js およびアプリ固有の .js ファイル) のローカルコピーは、Office アドインを開発する必要はありません (Office.js の CDN を参照すると、必要なファイルが実行時にダウンロードされます) が、ライブラリファイルのローカルコピーが必要な場合は、 [NuGet コマンドラインユーティリティ](https://docs.nuget.org/consume/installing-nuget) とコマンドを使用して `Install-Package Microsoft.Office.js` ダウンロードできます。

> [!NOTE]
> v1.1 アドイン マニフェストの XSD (XML スキーマ定義) のコピーの取得については、「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)」を参照してください。


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>最新のリリースを使用するようにプロジェクトの Office JavaScript API ライブラリファイルを更新する

1. テキスト エディターまたは IDE でアドインの HTML ページを開きます。

2. アドインの HTML ページの **head** タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された OFFICE JavaScript API ライブラリを次のように参照します。

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
> アドインマニフェストスキーマのバージョンを1.1 に更新した後で、 **機能** 要素と **機能** 要素を削除し、それらを [Hosts](../reference/manifest/hosts.md) 要素と [Host](../reference/manifest/host.md) 要素、または [要件と要件要素](specify-office-hosts-and-api-requirements.md)のいずれかに置き換える必要があります。

## <a name="see-also"></a>関連項目

- [Office アプリケーションと API の要件を指定する](specify-office-hosts-and-api-requirements.md) ]
- [Office JavaScript API について](understanding-the-javascript-api-for-office.md)
- [Office の JavaScript API](../reference/javascript-api-for-office.md)
- [Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)
