---
title: Office アドインの Fabric Core
description: このアドインで Fabric Core および Fabric UI コンポーネントを使用する方法のOffice説明します。
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cd534809bb443134e2df06de478e8283a3452aac
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150120"
---
# <a name="fabric-core-in-office-add-ins"></a>Office アドインの Fabric Core

Fabric Core は、CSS クラスと SASS mixins のオープンソース コレクションであり、このコレクションは、アドイン以外のアドインで使用React Officeです。Fabric Core には、アイコン、色、書体、グリッドFluent UI デザイン言語の基本的な要素が含まれています。 Fabric Core はフレームワークに依存しないので、任意の単一ページ アプリケーションまたは任意のサーバー側 Web UI フレームワークで使用できます。 (これは、歴史的な理由から、"Fluent コア" の代わりに "Fabric Core" と呼ばられています。

アドインの UI が Reactベースでない場合は、一連の非カスタム コンポーネントReactできます。 「USE [Office UI Fabric JS コンポーネント」を参照してください](#use-office-ui-fabric-js-components)。

> [!NOTE]
> この記事では、アドインのコンテキストでの Fabric Core のOffice説明します。ただし、さまざまなアプリや拡張機能でもMicrosoft 365使用されます。 詳細については[、「Fabric Core」](https://developer.microsoft.com/fluentui#/get-started/web#fabric-core)および「Open source repo Office UI Fabric [Core」を参照してください](https://github.com/OfficeDev/office-ui-fabric-core)。

## <a name="use-fabric-core-icons-fonts-colors"></a>Fabric Core を使用する: アイコン、フォント、色

1. ページの HTML に CDN 参照を追加します。  

    ```html
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css">
    ```

2. Fabric Core のアイコンとフォントを使用します。

    Fabric Core アイコンを使用するには、ページに "i" 要素を含め、適切なクラスを参照します。 アイコンのサイズは、フォント サイズを変更することで制御できます。 たとえば、次のコードは、themePrimary (#0078d7) 色を使用する特大の表アイコンを作成する方法を示しています。

    ```html
    <i class="ms-Icon ms-font-xl ms-Icon--Table ms-fontColor-themePrimary"></i>
    ```

    詳細な手順については、「UI[アイコンFluentを参照してください](https://developer.microsoft.com/fluentui#/styles/web/icons)。 Fabric Core で使用可能なアイコンを見つけるには、そのページの検索機能を使用します。 アドインで使用するアイコンを検索するときには、アイコン名の先頭に `ms-Icon--` を追加してください。

    Fabric Core で使用できるフォント サイズと色の詳細については、「色」の [「Typography」](https://developer.microsoft.com/fluentui#/styles/web/typography) および **「Colors」** の目次を参照 [してください](https://developer.microsoft.com/fluentui#/styles/web/colors)。

例については、この記事の [後半の「サンプル](#samples) 」に含まれています。

## <a name="use-office-ui-fabric-js-components"></a>JS Office UI Fabricを使用する

非カスタム REACT のアドインでは[、Office UI Fabric JS](https://github.com/OfficeDev/office-ui-fabric-js)の多くのコンポーネント (ボタン、ダイアログ、ピッカーなど) を使用することもできます。 手順については、repo の readme を参照してください。

例については、この記事の [後半の「サンプル](#samples) 」に含まれています。

## <a name="samples"></a>サンプル

次のサンプル アドインでは、Fabric Core または JS コンポーネントOffice UI Fabric使用します。 これらのリポジトリの一部はアーカイブ済みであり、バグやセキュリティ修正プログラムで更新されなくなりましたが、それらを使用して Fabric Core および Fabric UI コンポーネントの使い方を学習できます。

- [Excelアドイン JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
- [Excelアドイン SalesLeads](https://github.com/OfficeDev/Excel-Add-in-SalesLeads)
- [Excelアドイン WoodGrove 経費の傾向](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends)
- [Excelコンテンツ アドイン Humongous Insurance](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance)
- [Officeアドイン ファブリック UI のサンプル](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample)
- [Office-Add-in-UX-Design-Patterns-Code](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Outlookアドイン GifMe](https://github.com/OfficeDev/Outlook-Add-in-GifMe)
- [PowerPointアドイン Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Word アドイン Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)
- [Word アドイン JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word アドイン MarkdownConversion](https://github.com/OfficeDev/Word-Add-in-MarkdownConversion)
