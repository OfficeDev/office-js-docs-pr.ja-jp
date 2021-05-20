---
title: マニフェスト ファイルのランタイム
description: ランタイム要素は、アドインのランタイムを指定します。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 80336674c6d954bb9e0c6892feb41cb2f03c5859
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555298"
---
# <a name="runtimes-element"></a>ランタイム要素

アドインのランタイムを指定します。 要素の子 [`<Host>`](host.md) 。

> [!NOTE]
> WindowsでOfficeで実行する場合、マニフェストに要素を持つアドイン `<Runtimes>` は、必ずしも他の方法と同じ WebView コントロールで実行されるとは限りません。 WindowsとOfficeのバージョンが通常使用される webview コントロールを決定する方法の詳細については、「Office[アドインで使用されるブラウザー](../../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。WebView2 (Chromium ベース) でMicrosoft Edgeを使用する場合、アドインは要素を持つかどうかにかかわらず、そのブラウザーを使用します `<Runtimes>` 。 ただし、これらの条件が満たされない場合、要素を含むアドインでは `<Runtimes>` 、WindowsやMicrosoft 365のバージョンに関係なく、常に Internet Explorer 11 が使用されます。

**アドインの種類:** 作業ウィンドウ,メール

[!include[Runtimes support](../../includes/runtimes-note.md)]

## <a name="syntax"></a>構文

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>含まれる場所

[Host](host.md)

## <a name="child-elements"></a>子要素

|  要素 |  必須  |  説明  |
|:-----|:-----|:-----|
| [ランタイム](runtime.md) | はい |  アドインのランタイム。 **重要**: 現在、定義できる要素は 1 つだけです `<Runtime>` 。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)
