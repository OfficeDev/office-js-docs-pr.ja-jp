---
title: マニフェスト ファイル内のランタイム
description: Runtimes 要素は、アドインのランタイムを指定します。
ms.date: 04/16/2021
localization_priority: Normal
ms.openlocfilehash: 8f4a602c05b9af7bde9f644ef40b61a214e66cd5
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917087"
---
# <a name="runtimes-element"></a>Runtimes 要素

アドインのランタイムを指定します。 要素の [`<Host>`](host.md) 子。

> [!NOTE]
> Windows 上Officeで実行する場合、マニフェストに要素を持つアドインは、それ以外の場合と同じ Web ビュー コントロールで必ずしも `<Runtimes>` 実行されるとは限りません。 Windows および Officeのバージョンでどの Web ビュー コントロールが通常使用されるのかを決定する方法の詳細については、「Office アドインで使用されるブラウザー」を [参照してください](../../concepts/browsers-used-by-office-web-add-ins.md)。WebView2 で Microsoft Edge を使用する条件 (クロムベース) が満たされている場合、アドインはそのブラウザーが要素を持っているかどうかに応じ、そのブラウザーを使用 `<Runtimes>` します。 ただし、これらの条件が満たされない場合、要素を持つアドインは、Windows または Microsoft 365 バージョンに関係なく、常に Internet Explorer `<Runtimes>` 11 を使用します。

**アドインの種類:** 作業ウィンドウ, メール

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
| [ランタイム](runtime.md) | はい |  アドインのランタイム。 |

## <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
- [Office アドインを構成して共有 JavaScript ランタイムを使用する](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [イベント ベースのライセンス認証用に Outlook アドインを構成する](../../outlook/autolaunch.md)
