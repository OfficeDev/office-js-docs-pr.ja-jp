---
title: アドインでサポートされていないウィンドウ Officeオブジェクト
description: この記事では、アドインで動作しない一部のウィンドウ ランタイム オブジェクトOfficeします。
ms.date: 07/10/2020
ms.localizationpriority: medium
ms.openlocfilehash: 65cdd4d53dcbcdea75f7eeec39300e4eaee132ac
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154738"
---
# <a name="window-objects-that-are-unsupported-in-office-add-ins"></a>アドインでサポートされていないウィンドウ Officeオブジェクト

一部のバージョンの WindowsおよびOfficeアドインは、11 のランタイムInternet Explorer実行します。 (詳細については、「[アドインで使用Officeブラウザー」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。グローバル オブジェクトの一部のプロパティまたはサブプロパティは、11 ではInternet Explorer `window` されません。 これらのプロパティは、アドインで無効にされ、アドインが使用しているブラウザーに関係なく、アドインがすべてのユーザーに一貫したエクスペリエンスを提供します。 また、AngularJS が正しく読み込まれるのにも役立ちます。

無効になっているプロパティの一覧を次に示します。 リストは進行中の作業です。 アドインで動作しないその他のプロパティが見つかった場合は、以下のフィードバック ツールを使用 `window` して説明してください。

- `window.history.pushState`
- `window.history.replaceState`

## <a name="see-also"></a>関連項目

- [Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)