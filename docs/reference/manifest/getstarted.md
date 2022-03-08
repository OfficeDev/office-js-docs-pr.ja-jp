---
title: マニフェスト ファイルの GetStarted 要素
description: Word、Excel、PowerPoint、およびアドインにアドインがインストールされている場合に表示される吹き出しで情報をOneNote。
ms.date: 02/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 493526c3ad4a8486b76a18ccf23c64720a359784
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340996"
---
# <a name="getstarted-element"></a>GetStarted 要素

Word、Excel、PowerPoint、およびアドインにアドインがインストールされている場合に表示される吹き出しで情報をOneNote。 **GetStarted 要素** は [DesktopFormFactor の子要素です](desktopformfactor.md)。 **GetStarted 要素を省略** すると、代わりに [DisplayName](displayname.md) 要素と Description 要素の値 [が](description.md)使用されます。

**アドインの種類:** 作業ウィンドウ

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0

詳細については、「Version [overrides in the manifest」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

**次の要件セットに関連付けられている**。

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md)

## <a name="child-elements"></a>子要素

| 要素                       | 必須 | 説明                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | はい      | 吹き出しの一番上に使用するタイトル。     |
| [説明](#description)   | はい      | 吹き出しの説明/本文の内容。|
| [LearnMoreUrl](#learnmoreurl) | はい       | アドインの詳細を説明するページの URL。   |

### <a name="title"></a>タイトル 

必須。 吹き出しの一番上に使用するタイトル。 **resid 属性** は、[リソース] セクションの **ShortStrings** 要素の有効な ID [](resources.md) を参照し、32 文字以内で指定できます。

### <a name="description"></a>説明

必須。 吹き出しの説明/本文の内容。 **resid 属性** は、[リソース] セクションの **LongStrings** 要素の有効な ID [](resources.md) を参照し、32 文字以内で指定できます。

### <a name="learnmoreurl"></a>LearnMoreUrl

必須。 ユーザーがアドインの詳細を参照できるページの URL。 **resid 属性** は、[リソース] セクションの **Urls** 要素の有効な ID [](resources.md) を参照し、32 文字以内で指定できます。

> [!NOTE]
> **LearnMoreUrl** は現在、Word、Excel、または PowerPoint のクライアントではレンダリングされません。 これが利用可能になったときに URL がレンダリングされるよう、すべてのクライアントにこの URL を追加することをお勧めします。 

## <a name="see-also"></a>関連項目

次のコード サンプルでは、 **GetStarted 要素を使用** します。

* [テーブルとグラフの書式設定を操作するための Excel Web アドイン](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [Word アドインの JavaScript SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [PowerPoint アドインで Microsoft Graph を使用して Excel グラフを挿入する](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
