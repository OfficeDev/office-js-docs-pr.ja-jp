---
title: Office のバージョンと要件セット
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: 946068681070e8f482a8501639f40de4e13a7513
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325137"
---
# <a name="office-versions-and-requirement-sets"></a>Office のバージョンと要件セット

Office にはプラットフォームやバージョンが異なるものが数多くあり、それらは Office JavaScript API (Office.js) に含まれる API をすべてサポートしているわけではありません。 ユーザーがインストールしている Office のバージョンを制御できない場合があります。  このような状況に対処するため、Office アドインで必要な機能を Office ホストがサポートしているかどうかを判別するのに役立つ要件セットと呼ばれるシステムが用意されています。 

> [!NOTE]
> - Office は、Windows、ブラウザー、Mac、iPad などの複数のプラットフォームで実行されます。
> - Office ホストの例は、Excel、Word、PowerPoint、Outlook、OneNote などの Office 製品です。  
> - 要件セットとは、`ExcelApi 1.5` や `WordApi 1.3` などの、API メンバーの名前付きグループです。  

## <a name="how-to-check-your-office-version"></a>Office のバージョンを確認する方法

使用している Office のバージョンを特定するには、Office アプリケーション内で **[ファイル]** メニューを選択し、**[アカウント]** を選択します。 Office のバージョンは **[製品情報]** セクションに表示されます。 たとえば、次のスクリーン ショットは、Office のバージョンが 1802 (ビルド 9026.1000) であることを示しています。

![Office のバージョン確認](../images/office-version.png)

## <a name="office-requirement-sets-availability"></a>Office 要件セットの可用性

Office アドインは API 要件セットを使用して、使用する必要のある API メンバーを Office ホストがサポートしているかどうかを判別できます。 要件セットのサポートは、Office ホストと Office ホストのバージョンによって異なります (前のセクションを参照してください)。

一部の Office ホストには独自の API 要件セットがあります。 たとえば、Excel API の最初の要件セットは `ExcelApi 1.1` で、Word API の最初の要件セットは `WordApi 1.1` でした。 それ以降、追加の API 機能を提供するため、複数の新しい ExcelApi 要件セットと WordApi 要件セットが追加されています。

さらに、アドイン コマンド (リボン機能拡張) やダイアログ ボックスを起動する機能 (ダイアログ API) など、他の機能が共通 API に追加されました。 アドイン コマンドやダイアログ API の要件セットは、さまざまな Office ホストで共有されている API セットの例です。

アドインは、そのアドインが動作している Office ホストのバージョンでサポートしている要件セットにある API のみを使用できます。 特定の Office ホストのバージョンで使用できる要件セットを正確に確認するには、ホスト固有の要件セットに関する次の記事を参照してください。

- [Excel JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets) (ExcelApi、[CustomFunctionsRuntime](../excel/custom-functions-architecture.md))
- [Word JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets) (WordApi)
- [OneNote JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets) (OneNoteApi)
- [PowerPoint JavaScript API 要件セット](/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets) (PowerPointApi)
- [Outlook API 要件セットについて](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) (Mailbox)

一部の要件セットには、どの Office ホストでも使用できる API が含まれています。 これらの要件のセットの詳細については、次の記事を参照してください。

- [Office の共通要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [アドイン コマンドの要件セット](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [ダイアログ API の要件セット](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [Identity API の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)

`ExcelApi 1.1` の "1.1" など、要件セットのバージョン番号は Office ホストを基準にしています。 特定の要件セットのバージョン番号 (例: `ExcelApi 1.1`) は、Office.js のバージョン番号には対応しておらず、他の Office ホスト (Word、Outlook など) の要件セットにも対応していません。  Office ホストの要件セットがリリースされる早さや時期は、ホストによって異なります。 たとえば、`ExcelApi 1.5` の方が `WordApi 1.3` 要件セットより前にリリースされました。

Office JavaScript API ライブラリ (Office.js) には、現在利用可能なすべての要件セットが含まれています。 `ExcelApi 1.3` や `WordApi 1.3` のような要件セットは存在しますが、`Office.js 1.3` のような要件セットは存在しません。 Office.js の最新リリースは、コンテンツ配信ネットワーク (CDN) 経由で配信される単一の Office エンドポイントとして維持されます。 バージョン管理や下位互換性の処理方法など、Office.js CDN に関する詳細については、「[Office JavaScript API について](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)」を参照してください。

## <a name="specify-office-hosts-and-requirement-sets"></a>Office ホストと要件セットを指定する

アドインに必要となる Office ホストと要件セットは、さまざまな方法で指定できます。  詳細については、「[Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office の最新バージョンをインストールする](/office/dev/add-ins/develop/install-latest-office-version)
- [Office 365 ProPlus 更新プログラムのチャネルの概要](/deployoffice/overview-of-update-channels-for-office-365-proplus)
- [Office 365 で Office を最大限に活用する](https://products.office.com/compare-all-microsoft-office-products?tab=2)
