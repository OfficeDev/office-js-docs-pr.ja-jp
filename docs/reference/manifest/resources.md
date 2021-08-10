---
title: マニフェスト ファイルの Resources 要素
description: Resources 要素には、VersionOverrides ノードのアイコン、文字列、URL が含まれます。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 9fbf28790b717389bf7248babcb69f28d4226b77c3139415fb14bbd712690074
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094146"
---
# <a name="resources-element"></a>リソース要素

[VersionOverrides](versionoverrides.md) ノードのアイコン、文字列、および URL が含まれます。 マニフェスト要素によりリソースが指定されます。リソースの **id** を使用します。 それにより、特にリソースにさまざまなロケールのバージョンがあるとき、マニフェストのサイズが管理できる大きさに抑えられます。 **id** はマニフェスト内で一意にする必要があり、最大 32 文字を使用できます。

各リソースは、特定のロケールに異なるリソースを定義する 1 つ以上の **Override** 子要素を持つことができます。

## <a name="child-elements"></a>子要素

|  要素 |  型  |  説明  |
|:-----|:-----|:-----|
|  [Images](#images)            |  image   |  アイコンの画像への HTTPS URL を指定します。 |
|  **Urls**                |  url     |  HTTPS URL の場所を指定します。 URL の長さは最大で 2048 文字です。 |
|  **ShortStrings** |  string  |  **Label** 要素と **Title** 要素のテキスト。 各 **String** には、最大 125 文字を使用できます。|
|  **LongStrings**  |  string  | **Description** 属性のテキスト。 各 **String** には、最大 250 文字を使用できます。|

> [!NOTE]
> **Image** 要素と **Url** 要素のすべての URL で Secure Sockets Layer (SSL) を使用する必要があります。

### <a name="images"></a>画像

各アイコンには、3 つの必須サイズごとに 1 つの **3** つの Images 要素が必要です。

- 16x16
- 32x32
- 80x80

次の追加サイズもサポートされますが、必須ではありません。

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT]
>
> - この画像がアドインの代表的なアイコンである場合は、「Create effective listings in [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office サイズと他の要件」を参照してください。
> - Outlook では、パフォーマンス向上のために画像リソースをキャッシュする機能が必要です。 このため、画像リソースをホストするサーバーは、どんな CACHE-CONTROL ディレクティブも応答ヘッダーに追加することはできません。 これは、Outlook が汎用の画像や既定の画像を自動的に代用する原因になります。

## <a name="resources-examples"></a>リソースの例

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
