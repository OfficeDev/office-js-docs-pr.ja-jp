---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットの詳細
ms.date: 11/19/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: e583c9832f04e17dc1c82d38d056fe2749888a77
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757493"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API のオンラインのみの要件セット

`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。 この要件セットの Api は、web ホスト上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。 `ExcelApiOnline`他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。

`ExcelApiOnline`要件セットの api がすべてのプラットフォームでサポートされている場合は、次にリリースされる`ExcelApi 1.[NEXT]`要件セット () に追加されます。 新しい要件が公開されると、これらの Api はから`ExcelApiOnline`削除されます。 この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。

> [!IMPORTANT]
> `ExcelApiOnline`は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1`は、オンライン専用 Api の唯一のバージョンです。 これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。

## <a name="recommended-usage"></a>推奨される使用法

Api `ExcelApiOnline`は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。 これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

クロスプラットフォームの要件セットに含まれる API は、 `isSetSupported`チェックを削除または編集する必要があります。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行うときは、これらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェストでライセンス認証`ExcelApiOnline 1.1`の要件として指定することはできません。 [Set 要素](../manifest/set.md)で使用する有効な値ではありません。

## <a name="api-list"></a>API リスト

現時点では、オンライン専用の Api はありません。 新機能が web 上の Excel に追加され、Office JavaScript Api によってサポートされるようになると、もう一度確認してください。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online)
- [Excel JavaScript プレビュー API](./excel-preview-apis.md)
- [Excel JavaScript API の要件セット](./excel-api-requirement-sets.md)