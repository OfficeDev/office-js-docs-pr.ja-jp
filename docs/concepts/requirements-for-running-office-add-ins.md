---
title: Office アドインを実行するための要件
description: エンド ユーザーがアドインで実行する必要があるクライアント要件とサーバー要件Office説明します。
ms.date: 09/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: b39af2b381bc6dd29df2f1925ca5cbf67740e4a8
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990559"
---
# <a name="requirements-for-running-office-add-ins"></a>Office アドインを実行するための要件

この記事では、Office アドインを実行するためのソフトウェアとデバイスの要件について説明します。

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Office アドインが現在サポートされている場所の詳細なビューについては[、「Office](../overview/office-add-in-availability.md)クライアント アプリケーションとプラットフォームの可用性」を参照Officeしてください。

## <a name="server-requirements"></a>サーバーの要件

Office アドインをインストールおよび実行できるようにするには、まずアドインの UI とコードのマニフェストと Web ページ ファイルを、適切なサーバーの場所に展開する必要があります。

すべての種類のアドイン (コンテンツ、Outlook、作業ウィンドウの、アドインとアドイン コマンド) で、アドインの Web ページ ファイルを Web サーバーや [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md) などの Web ホスティング サービスに展開する必要があります。

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Visual Studio でアドインを開発およびデバッグする際、Visual Studio は IIS Express を使用してアドインの Web ページ ファイルをローカルで展開および実行するので、追加の Web サーバーは必要ありません。

コンテンツ アドインと作業ウィンドウ アドインの場合、サポートされている Office クライアント アプリケーション (Excel、PowerPoint、Project、または Word) では、アドインの XML マニフェスト ファイルを[](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)アップロードするために SharePoint のアプリ カタログも必要か、統合アプリを使用してアドインを展開する必要があります。 [](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)

Outlook アドインをテストして実行するには、ユーザーの Outlook 電子メール アカウントが Exchange 2013 以降に存在し、Microsoft 365、Exchange Online、またはオンプレミスインストールを通じて使用できる必要があります。 ユーザーまたは管理者は、サーバー上に Outlook アドインのマニフェスト ファイルをインストールします。

> [!NOTE]
> Outlook の POP および IMAP 電子メール アカウントは、Office アドインをサポートしていません。

## <a name="client-requirements-windows-desktop-and-tablet"></a>クライアントの要件: Windows デスクトップおよびタブレット

Windows ベースのデスクトップ、ラップトップ、またはタブレット デバイスで実行されるサポートされている Office デスクトップ クライアントまたは Web クライアント用の Office アドインを開発するには、次のソフトウェアが必要です。

- Windows x86 および x64 デスクトップおよび Surface Pro などのタブレット:
  - Windows 7 以降のバージョンで実行している Office 2013 以降のバージョンの、32 ビットまたは 64 ビット バージョン。
  - Excel 2013、Outlook 2013、PowerPoint 2013、Project Professional 2013、Project 2013 SP1、Word 2013、またはそれ以降の Office クライアントのバージョン (特にこれらの Office デスクトップ クライアントを対象として Office アドインをテストまたは実行する場合)。Office デスクトップ クライアントはオンプレミスでインストールすることも、クイック実行によってクライアント コンピューターにインストールすることもできます。

  有効なサブスクリプションをMicrosoft 365、Office クライアントにアクセスできない場合は、最新バージョンの Office を[ダウンロードしてインストールできます](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658)。

- Internet Explorer 11 または Microsoft Edge (Windows および Office のバージョンによる) がインストールされている必要がありますが、既定のブラウザーである必要はありません。 Office アドインをサポートするために、ホストとして動作する Office のクライアントは、Internet Explorer 11 または Microsoft Edge に組み込まれているブラウザー コンポーネントを使用します。 詳細については、「[Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)」を参照してください。

  > [!NOTE]
  > Office Web アドインが機能するためには、Internet Explorer のセキュリティ強化の構成 (ESC) がオフになっている必要があります。 アドインを開発する際に Windows Server コンピューターをクライアントとして使用する場合は、Windows Server では既定で ESC がオンになっていることに注意してください。

- 既定のブラウザーとして次のいずれか: Internet Explorer 11、または Microsoft Edge、Chrome、Firefox、Safari (Mac OS) の最新バージョンのうちいずれか。
- メモ帳などの HTML および JavaScript エディター、[Visual Studio および Microsoft Developer Tools](https://www.visualstudio.com/features/office-tools-vs)、またはサードパーティの Web 開発ツール。

## <a name="client-requirements-os-x-desktop"></a>クライアントの要件: OS X デスクトップ

Outlookの一部として配布される Mac 上のMicrosoft 365は、Outlookアドインをサポートします。mac Outlook Outlook でアドインを実行する場合、Mac 自体で Outlook と同じ要件が必要です。オペレーティング システムは、少なくとも OS X v10.10 "Yosemite" である必要があります。 Mac 上の Outlook はレイアウト エンジンとして WebKit を使用して、アドイン ページを表示するので、追加のブラウザーの依存関係はありません。

次は、Office アドインをサポートする Mac 上の Office の最小クライアント バージョンです。

- Word バージョン 15.18 (160109)
- Excel バージョン 15.19 (160206)
- PowerPoint バージョン 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>クライアントの要件: Office Web クライアントと SharePoint のブラウザー サポート

ecMAScript 5.1、HTML5、CSS3 をサポートするブラウザー (Microsoft Edge、Chrome、Firefox、Safari (Mac OS) など、Internet Explorer を除くすべてのブラウザー。


## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>クライアントの要件: Windows 以外のスマートフォンおよびタブレット

特に、スマートフォンや Windows 以外のタブレット デバイス上のブラウザーで動作する Outlook の場合、Outlook アドインをテストおよび実行するのに以下のソフトウェアが必要です。


| Office アプリケーション | デバイス | オペレーティング システム | Exchange アカウント | モバイル ブラウザー |
|:-----|:-----|:-----|:-----|:-----|
|Android 上の Outlook|Android のタブレットとスマートフォン|Android 4.4 KitKat 以降|最新の更新プログラムのMicrosoft 365 Apps for businessまたはExchange Online|Android 用のネイティブ アプリ、ブラウザーは適用外|
|iOS 上の Outlook|iPad のタブレット、iPhone のスマート フォン|iOS 11 以降|最新の更新プログラムのMicrosoft 365 Apps for businessまたはExchange Online|iOS 用のネイティブ アプリ、ブラウザーは適用外|
|Outlook on the web|iPhone 4 以降、iPad 2 以降、iPod Touch 4 以降|iOS 5 以降|2013 Microsoft 365、Exchange Online、またはオンプレミスの 2013 以降Exchange Serverオン|Safari|

> [!NOTE]
> ネイティブ アプリの OWA for Android、OWA for iPad、および OWA for iPhone は[廃止](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b)され、Outlook アドインのテストには不要になり、利用もできなくなりました。


## <a name="see-also"></a>関連項目

- [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
- [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性](../overview/office-add-in-availability.md)
- [Office アドインによって使用されるブラウザー](browsers-used-by-office-web-add-ins.md)
