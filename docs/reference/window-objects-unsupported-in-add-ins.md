---
title: Office アドインでサポートされていない Window オブジェクト
description: この記事では、Office アドインでは動作しない window ランタイムオブジェクトの一部について説明します。
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: d2560748841bd1e2a7708b25a8e51133563d1534
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160506"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>Office アドインでサポートされていない Window オブジェクト

Windows および Office の一部のバージョンでは、アドインは Internet Explorer 11 ランタイムで実行されます。 (詳細については、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください)。グローバルオブジェクトの一部のプロパティまたはサブプロパティは、 `window` Internet Explorer 11 ではサポートされていません。 アドインで使用されているブラウザーに関係なく、すべてのユーザーに一貫した機能を提供するために、これらのプロパティはアドインで無効になっています。 これは、AngularJS が適切に読み込まれるのにも役に立ちます。

無効にされたプロパティの一覧を次に示します。 リストは処理中です。 アドインで機能しない他のプロパティが見つかった場合は、 `window` 次のフィードバックツールを使用してご確認ください。

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>関連項目

- [Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)