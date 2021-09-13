---
title: マニフェスト ファイルの DesktopSettings 要素
description: メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。
ms.date: 04/09/2020
ms.localizationpriority: medium
ms.openlocfilehash: 9393871e56f686b710ffd0031f93e776f362a89d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151615"
---
# <a name="desktopsettings-element"></a>DesktopSettings 要素

メール アドインがデスクトップ コンピューターで使用されるときに適用されるソースの場所と制御の設定を指定します。

> [!IMPORTANT]
> この要素は、従来の Outlook on the web (通常はオンプレミスの Exchange サーバーの古いバージョンに接続されている) および Outlook `DesktopSettings` 2013 Windows。

**アドインの種類:** メール

## <a name="syntax"></a>構文

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a>含まれる場所

[Form](form.md)
