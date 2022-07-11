---
title: ネットワーク共有からテストするための Office アドインをサイドロードする
description: ネットワーク共有からテストするために Office アドインをサイドロードする方法について説明します。
ms.date: 05/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: 87bdeb6cbd33bcd9b1828c7afa0a9f879d4c05e4
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712777"
---
# <a name="sideload-office-add-ins-for-testing-from-a-network-share"></a>ネットワーク共有からテストするための Office アドインをサイドロードする

Windows 上の Office クライアントで Office アドインをテストするには、マニフェストをネットワーク ファイル共有に発行します (以下の手順を参照)。 このデプロイ オプションは、ローカル ホストでの開発とテストを完了し、ローカル以外のサーバーまたはクラウド アカウントからアドインをテストする場合に使用することを目的としています。

> [!IMPORTANT]
> 運用アドインでは、ネットワーク共有によるデプロイはサポートされていません。このメソッドには、次の制限事項があります。
>
> - アドインは Windows コンピューターにのみインストールできます。
> - 新しいバージョンのアドインでリボンが変更された場合 (カスタム タブやカスタム ボタンを追加するなど) は、各ユーザーがアドインを再インストールする必要があります。

> [!NOTE]
> アドイン プロジェクトが十分に新しい [Office 用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md) バージョンで作成されている場合、アドインは `npm start` を実行すると自動的に Office デスクトップ クライアントにサイドロードします。

この記事は、Word、Excel、PowerPoint、および Project アドインのテストにのみ適用され、Windows でのみ適用されます。 別のプラットフォームでテストする場合、または Outlook アドインをテストする場合は、次のいずれかのトピックを参照してアドインをサイドロードします。

- [テスト用に Office on the web で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [テスト用の Mac 上の Office アドインをサイドロードする](sideload-an-office-add-in-on-mac.md)
- [iPad でテスト用の Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)

次のビデオでは、共有フォルダー カタログを使用して Office on the web またはデスクトップでアドインをサイドロードする手順について説明します。  

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="share-a-folder"></a>フォルダーの共有

1. アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。

1. 共有フォルダー カタログとして使用するフォルダーのコンテキスト メニューを開き (フォルダーを右クリック)、[**プロパティ**] を選択します。

1. [**プロパティ**] ダイアログ ボックス内で [**共有**] タブを選択し、[**共有**] ボタンを選択します。

    ![[共有] タブと [共有] ボタンが強調表示された [フォルダーのプロパティ] ダイアログ。](../images/sideload-windows-properties-dialog.png)

1. [**ネットワーク アクセス**] ダイアログ ウィンドウで自分自身とアドインを共有する相手のユーザーまたはグループを追加します。 最低でも、フォルダーへの **読み取り/書き込み** アクセス許可が必要です。 共有する相手の選択が完了したら、[**共有**] ボタンを選択します。

1. 「**ユーザーのフォルダーは共有されています**」という確認メッセージが表示されたら、フォルダー名のすぐ後に表示される完全なネットワーク パスを書き留めます。 (この記事の次のセクションで説明する通り、[共有フォルダーを信頼できるカタログとして指定する](#specify-the-shared-folder-as-a-trusted-catalog)際に、このネットワーク パスを [**カタログの URL**] として入力する必要があります。) [**完了**] を選択して [**ネットワーク アクセス**] ダイアログ ウィンドウを閉じます。

   ![共有パスが強調表示された [ネットワーク アクセス] ダイアログ。](../images/sideload-windows-network-access-dialog.png)

1. [**閉じる**] を選択して、[**プロパティ**] ダイアログ ウィンドウを閉じます。

## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>共有フォルダーを信頼できるカタログとして指定する

### <a name="configure-the-trust-manually"></a>信頼を手動で構成する

1. Excel、Word、PowerPoint、または Project で新しいドキュメントを開きます。

1. [**ファイル**] タブを選択し、[**オプション**] を選択します。

1. [**セキュリティ センター**] を選択し、[**セキュリティ センターの設定**] ボタンを選択します。

1. [**信頼されているアドイン カタログ**] を選びます。

1. [**カタログの URL**] ボックスで、先ほど [共有](#share-a-folder)したフォルダーの完全なネットワーク パスを入力します。 フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。

    ![[共有] タブとネットワーク パスが強調表示された [フォルダーのプロパティ] ダイアログ。](../images/sideload-windows-properties-dialog-2.png)

1. [**カタロ URL**] ボックスにフォルダーの完全なネットワーク パスを入力したら、[**カタログの追加**] を選択します。

1. 新しく追加されたアイテムの [**メニューに表示する**] チェック ボックスをオンにし、[**OK**] を選択して [**セキュリティ センター** ] ダイアログ ウィンドウを閉じます。 

    ![カタログが選択された [セキュリティ センター] ダイアログ。](../images/sideload-windows-trust-center-dialog.png)

1. **[OK]** ボタンを選択して **、[オプション]** ダイアログ ウィンドウを閉じます。

1. Office アプリケーションを閉じてからもう一度開くと変更内容が有効になります。

### <a name="configure-the-trust-with-a-registry-script"></a>レジストリ スクリプトを使用して信頼を構成する

1. テキスト エディターで、TrustNetworkShareCatalog.reg という名前のファイルを作成します。

1. 次のコンテンツをファイルに追加します。

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-random-GUID-here-}]
    "Id"="{-random-GUID-here-}"
    "Url"="\\\\-share-\\-folder-"
    "Flags"=dword:00000001
    ```

1. [GUID ジェネレーター](https://guidgenerator.com/)など、多数のオンライン GUID 生成ツールのいずれかを使用してランダムな GUID を生成し、TrustNetworkShareCatalog.reg ファイル内で *両方の場所* の文字列「-random-GUID-here-」を GUID に置き換えます。 (引用符 `{}` 記号は残しておく必要があります)。

1. `Url` 値を、以前[共有](#share-a-folder)したフォルダーへの完全なネットワーク パスに置き換えます。 (URL の `\` 文字は 2 倍にする必要があります。) フォルダーを共有した際に完全なネットワーク パスを書き留めておかなかった場合は、次のスクリーン ショットに示されるように、フォルダーの [**プロパティ**] ダイアログ ウィンドウから取得できます。

    ![[共有] タブとネットワーク パスが強調表示された [フォルダーのプロパティ] ダイアログ。](../images/sideload-windows-properties-dialog-2.png)

1. ファイルは、次のようになります。 ファイルを保存します。

    ```text
    Windows Registry Editor Version 5.00

    [HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{01234567-89ab-cedf-0123-456789abcedf}]
    "Id"="{01234567-89ab-cedf-0123-456789abcedf}"
    "Url"="\\\\TestServer\\OfficeAddinManifests"
    "Flags"=dword:00000001
    ```

1. *すべて* の Office アプリケーションを閉じます。

1. ダブルクリックするなど、実行可能ファイルと同様に TrustNetworkShareCatalog.reg 実行します。

## <a name="sideload-your-add-in"></a>アドインのサイドロード

1. テストするアドインのマニフェスト XML ファイルを共有フォルダー カタログに置きます。 なお、Web アプリケーション自体を Web サーバーに展開します。 マニフェスト ファイルの要素で **\<SourceLocation\>** URL を指定してください。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

    > [!NOTE]
    > Visual Studio プロジェクトの場合は、フォルダー内のプロジェクトによってビルドされたマニフェストを `{projectfolder}\bin\Debug\OfficeAppManifests` 使用します。

1. Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。 Projectで、リボンの [**Project**]タブの [**個人用アドイン**] を選択します。

1. **[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。

1. アドインの名前を選び、**[追加]** を選択して、アドインを挿入します。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

以前にサイドロードされたアドインを削除するには、コンピューター上の Office キャッシュをクリアします。 Windows でキャッシュをクリアする方法の詳細については、「 [Office キャッシュをクリアする」](clear-cache.md#clear-the-office-cache-on-windows)を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインのマニフェストを検証する](troubleshoot-manifest.md)
- [Office のキャッシュをクリアする](clear-cache.md)
- [Office アドインを発行する](../publish/publish.md)
