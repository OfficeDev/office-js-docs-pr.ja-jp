---
title: Office アドインを実行するための要件
description: エンド ユーザーがアドインOffice実行するために必要なクライアントとサーバーの要件について説明します。
ms.date: 06/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 06699e8a2c498eb6ad2f9832a8369beef5af4786
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091035"
---
# <a name="requirements-for-running-office-add-ins"></a>Office アドインを実行するための要件

この記事では、Office アドインを実行するためのソフトウェアとデバイスの要件について説明します。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Office アドインが現在サポートされている場所の概要については、「Office アドインの[クライアント アプリケーションとプラットフォームの可用性Office](/javascript/api/requirement-sets)参照してください。

## <a name="server-requirements"></a>サーバーの要件

Office アドインをインストールおよび実行できるようにするには、まずアドインの UI とコードのマニフェストと Web ページ ファイルを、適切なサーバーの場所に展開する必要があります。

すべての種類のアドイン (コンテンツ、Outlook、作業ウィンドウの、アドインとアドイン コマンド) で、アドインの Web ページ ファイルを Web サーバーや [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) などの Web ホスティング サービスに展開する必要があります。

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Visual Studio でアドインを開発およびデバッグする際、Visual Studio は IIS Express を使用してアドインの Web ページ ファイルをローカルで展開および実行するので、追加の Web サーバーは必要ありません。

コンテンツ および作業ウィンドウ アドインの場合、サポートされているOffice クライアント アプリケーション (Excel、PowerPoint、Project、Word) では、アドインの XML マニフェスト ファイルをアップロードするためにSharePointの[アプリ カタログ](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)も必要です。また、[統合アプリ](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)を使用してアドインを展開する必要があります。

Outlook アドインをテストして実行するには、ユーザーのOutlook電子メール アカウントが 2013 以降Exchangeに存在する必要があります。これは、Microsoft 365、Exchange Online、またはオンプレミスのインストールを通じて使用できます。 ユーザーまたは管理者は、サーバー上に Outlook アドインのマニフェスト ファイルをインストールします。

> [!NOTE]
> Outlook の POP および IMAP 電子メール アカウントは、Office アドインをサポートしていません。

## <a name="client-requirements-windows-desktop-and-tablet"></a>クライアントの要件: Windows デスクトップおよびタブレット

Windows ベースのデスクトップ、ノート PC、またはTablet PC デバイスで実行される、サポートされているOffice デスクトップ クライアントまたは Web クライアント用のOffice アドインを開発するには、次のソフトウェアが必要です。

- Windows x86 および x64 デスクトップおよび Surface Pro などのタブレット:
  - Windows 7 以降のバージョンで実行している Office 2013 以降のバージョンの、32 ビットまたは 64 ビット バージョン。
  - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013、またはそれ以降の Office クライアントのバージョン (特にこれらの Office デスクトップ クライアントを対象として Office アドインをテストまたは実行する場合)。Office デスクトップ クライアントはオンプレミスでインストールすることも、クイック実行によってクライアント コンピューターにインストールすることもできます。

  有効なMicrosoft 365 サブスクリプションがあり、Office クライアントにアクセスできない場合は、[最新バージョンのOfficeをダウンロードしてインストール](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658)できます。

- Microsoft Edgeをインストールする必要がありますが、既定のブラウザーである必要はありません。 Office アドインをサポートするために、ホストとして機能するOffice クライアントは、Microsoft Edgeの一部であるブラウザー コンポーネントを使用します。

  > [!NOTE]
  >
  > - 厳密に言えば、Internet Explorer 11 がインストールされているが、Microsoft Edgeされていないコンピューターでアドインを開発することは可能です。 ただし、IE は、特定の古いバージョンのWindowsとOfficeバージョンの組み合わせでのみアドインを実行するために使用されます。 詳細については、「[Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)」を参照してください。 プライマリ アドイン開発環境のような古い環境を使用することはお勧めしません。 ただし、これらの古い組み合わせで動作しているアドインの顧客がいる可能性が高い場合は、Internet Explorer をサポートすることをお勧めします。 詳細については、「 [Internet Explorer 11 のサポート](../develop/support-ie-11.md)」を参照してください。
  > - Office Web アドインが機能するためには、Internet Explorer のセキュリティ強化の構成 (ESC) がオフになっている必要があります。 アドインを開発する際に Windows Server コンピューターをクライアントとして使用する場合は、Windows Server では既定で ESC がオンになっていることに注意してください。

- 既定のブラウザーとして次のいずれか: Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンのうちいずれか。
- [Visual Studio Code](https://code.visualstudio.com/)、[Visual Studio、Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)、Microsoft 以外の Web 開発ツールなどの HTML および JavaScript エディター。

## <a name="client-requirements-os-x-desktop"></a>クライアントの要件: OS X デスクトップ

Microsoft 365の一部として配布される Mac 上のOutlookは、Outlook アドインをサポートします。Mac でOutlookでOutlookアドインを実行する場合、Mac 自体のOutlookと同じ要件があります。オペレーティング システムは OS X v10.10 "Marketplace" 以上である必要があります。 Mac 上の Outlook はレイアウト エンジンとして WebKit を使用して、アドイン ページを表示するので、追加のブラウザーの依存関係はありません。

次は、Office アドインをサポートする Mac 上の Office の最小クライアント バージョンです。

- Word バージョン 15.18 (160109)
- Excel バージョン 15.19 (160206)
- PowerPoint バージョン 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>クライアントの要件: Office Web クライアントと SharePoint のブラウザー サポート

ECMAScript 5.1、HTML5、CSS3 をサポートする Internet Explorer を除くすべてのブラウザー (Microsoft Edge、Chrome、Firefox、Safari (Mac OS) など)。

## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>クライアント要件: 非WindowsスマートフォンとTablet PC

スマートフォンや非Windows Tablet PC デバイスで実行されるOutlookに特に、Outlook アドインのテストと実行には、次のソフトウェアが必要です。

| Office アプリケーション | デバイス | オペレーティング システム | Exchange アカウント | モバイル ブラウザー |
|:-----|:-----|:-----|:-----|:-----|
|Android 上の Outlook|- タブレットをAndroidする<br>- スマートフォンをAndroidする|- Android 4.4 KitKat 以降|Microsoft 365 Apps for businessまたはExchange Onlineの最新の更新時|ブラウザーは適用されません。 Androidにはネイティブ アプリを使用します。<sup>1</sup>|
|Outlook on iOS|- タブレットをiPadする<br>- スマートフォンをiPhoneする|- iOS 11 以降|Microsoft 365 Apps for businessまたはExchange Onlineの最新の更新時|ブラウザーは適用されません。 iOSにはネイティブ アプリを使用します。<sup>1</sup>|
|Outlook on the web (モダン)<sup>2</sup>|- iPad 2 以降<br>- タブレットをAndroidする |- iOS 5 以降<br>- Android 4.4 KitKat 以降|Microsoft 365では、Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>- Safari|
|Outlook on the web (クラシック)|- iPhone 4 以降<br>- iPad 2 以降<br>- iPod Touch 4 以降|- iOS 5 以降|オンプレミス Exchange Server 2013 以降<sup>3</sup>|- Safari|

> [!NOTE]
> <sup>Android</sup>の場合は 1 OWA、iPadの場合は OWA、ネイティブ アプリの場合は OWA iPhone[非推奨になりました](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b)。
>
> iPhoneおよびAndroidスマートフォンの <sup>2</sup> つの最新のOutlook on the webは、Outlook アドインのテストに必要な、または使用できなくなりました。
>
> <sup>3</sup> つのアドインは、Android、iOS、およびオンプレミスのExchange アカウントを持つ最新のモバイル Web のOutlookではサポートされていません。

> [!TIP]
> メールボックスのツールバーを確認することで、Web ブラウザーでの Outlook がモダンかクラシックかを区別できます。
>
> **モダン**
>
> ![Outlook ツールバー (モダン) の部分的なスクリーンショット。](../images/outlook-on-the-web-new-toolbar.png)
>
> **クラシック**
>
> ![Outlook ツールバー (クラシック) の部分的なスクリーンショット。](../images/outlook-on-the-web-classic-toolbar.png)

## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](/javascript/api/requirement-sets)
- [Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)
