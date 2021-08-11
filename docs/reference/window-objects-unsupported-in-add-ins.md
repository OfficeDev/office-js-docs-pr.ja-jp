---
title: アドインでサポートされていないウィンドウ Officeオブジェクト
description: この記事では、アドインで動作しない一部のウィンドウ ランタイム オブジェクトOfficeします。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 654e8e311069a616e2d8859a4f63b19d299609982fa68449b5529df489816cbf
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097385"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>アドインでサポートされていないウィンドウ Officeオブジェクト

一部のバージョンの WindowsおよびOfficeアドインは、11 のランタイムInternet Explorer実行します。 (詳細については、「[アドインで使用Officeブラウザー」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。グローバル オブジェクトの一部のプロパティまたはサブプロパティは、11 ではInternet Explorer `window` されません。 これらのプロパティは、アドインで無効にされ、アドインが使用しているブラウザーに関係なく、アドインがすべてのユーザーに一貫したエクスペリエンスを提供します。 また、AngularJS が正しく読み込まれるのにも役立ちます。

無効になっているプロパティの一覧を次に示します。 リストは進行中の作業です。 アドインで動作しないその他のプロパティが見つかった場合は、以下のフィードバック ツールを使用 `window` して説明してください。

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>関連項目

- [Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)