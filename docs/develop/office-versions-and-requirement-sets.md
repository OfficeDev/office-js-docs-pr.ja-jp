---
title: Office のバージョンと要件セット
description: ''
ms.date: 03/29/2018
ms.openlocfilehash: ac3ae4fa3eeca9cfbd56b15168fc39d67139680d
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505994"
---
# <a name="office-versions-and-requirement-sets"></a>Office のバージョンと要件セット

Office にはプラットフォームやバージョンが異なるものが数多くあり、それらすべてが Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。このような状況に対処するため、Office アドインで必要な機能を Office ホストがサポートしているかどうかを判別するのに役立つ要件セットと呼ばれるシステムが用意されています。 

> [!NOTE]
> - Office は、Office for Windows、Office Online、Office for Mac、Office for iPad を含む複数のプラットフォームで実行できます。  
> - Office ホストの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。  
> - 要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。  


## <a name="how-to-check-your-office-version"></a>Office のバージョンを確認する方法

使用している Office のバージョンを確認するには、Office アプリケーション内の **[ファイル]** メニューを選択し、**[アカウント]** を選択します。この Office のバージョンは、 **[製品情報]** セクションに表示されます。たとえば、次のスクリーンショットは Office バージョン 1802 (ビルド 9026.1000) を示しています

![Office のバージョン確認](../images/office-version-number-ui.jpg)


## <a name="office-requirement-sets-availability"></a>Office 要件セットの可用性

Office アドインは API 要件セットを使用して、使用する必要のある API メンバーを Office ホストがサポートしているかどうかを判別できます。要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります (前のセクションを参照してください)。

一部の Office ホストでは、独自の API 要件セットがあります。たとえば、Excel API の最初の要件セットは `ExcelApi 1.1` で、Word API の最初の要件セットは `WordApi 1.1`でした。それ以降、追加の機能を提供するため、複数の新しい ExcelApi 要件セットと WordApi 要件セットが追加されています。

さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が一般的な API に追加されました。アドイン コマンドやダイアログ API の要件セットは、さまざまな Office ホストで共有されている API セットの例です。

アドインは、そのアドインが動作している Office ホストのバージョンでサポートしている要件セットにある API のみを使用できます。特定の Office ホストのバージョンで使用できる要件セットを正確に確認するには、ホスト固有の要件セットに関する次の記事を参照してください。

- [Excel JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets?view=office-js) (ExcelApi)
- [Word JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets?view=office-js) (WordApi)
- [OneNote JavaScript API 要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets?view=office-js) (OneNoteApi)
- [Outlook API 要件セットについて](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets?view=office-js) (MailBox)

一部の要件セットには、どの Office ホストでも使用できる API が含まれています。これらの要件のセットの詳細については、次の記事を参照してください。

- [Office の共通要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets?view=office-js)
- [アドイン コマンドの要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets?view=office-js)
- [ダイアログ API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Identity API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)

`ExcelApi 1.1`の「1.1」など要件セットのバージョン番号は、Office ホストを基準としています。 特定の要件セットのバージョン番号 (たとえば、`ExcelApi 1.1`) は、Office.js や Office ホスト (たとえば Word、Outlook) の要件セットに対応しておらず、他の Office ホストの要件セットは、異なる時期にリリースされています。たとえば`ExcelApi 1.5` は`WordApi 1.3` 要件セットよりも前にリリースされました。

JavaScript API for Office ライブラリ (Office.js) には現在利用できるすべての要件セットが含まれています。要件セット `ExcelApi 1.3` や `WordApi 1.3` がある一方で、 `Office.js 1.3` 要件セットはありません。最新リリースの Office.js は、コンテンツ配信ネットワーク (CDN) を介して配信される単一 Office エンドポイントとして維持されます。バージョン管理と下位互換性の処理方法など、Office.js CDN に関する詳細は、「 [JavaScript API for Office を理解する](https://docs.microsoft.com/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。

## <a name="specify-office-hosts-and-requirement-sets"></a>Office ホストと要件セットを指定する

アドインに必要となる Office ホストと要件セットは、さまざまな方法で指定できます。詳細については、「 [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)」を参照してください。


## <a name="see-also"></a>関連項目

- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office の最新バージョンをインストールする](https://docs.microsoft.com/office/dev/add-ins/develop/install-latest-office-version)
- [Office 365 ProPlus 更新チャネルの概要](https://docs.microsoft.com/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Office 365 で Office を最大限に活用する](https://products.office.com/compare-all-microsoft-office-products?tab=2)
