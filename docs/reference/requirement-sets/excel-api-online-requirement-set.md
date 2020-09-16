---
title: Excel JavaScript API のオンラインのみの要件セット
description: ExcelApiOnline の要件セットに関する詳細。
ms.date: 09/15/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 29f5826ba2adbf18b79033b83254b046210015fe
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819806"
---
# <a name="excel-javascript-api-online-only-requirement-set"></a>Excel JavaScript API のオンラインのみの要件セット

`ExcelApiOnline`要件セットは、web 上の Excel でのみ使用可能な機能を含む特別な要件セットです。 この要件セットの Api は、web アプリケーション上の Excel の運用 Api (未提出の行動または構造上の変更による影響を受けない) と見なされます。 `ExcelApiOnline` 他のプラットフォーム (Windows、Mac、iOS) の場合は "preview" Api と見なされますが、これらのプラットフォームではサポートされていない場合があります。

要件セットの Api `ExcelApiOnline` がすべてのプラットフォームでサポートされている場合は、次にリリースされる要件セット () に追加され `ExcelApi 1.[NEXT]` ます。 新しい要件が公開されると、これらの Api はから削除され `ExcelApiOnline` ます。 この点は、プレビューからリリースに移行する API と同様に、昇格プロセスと考えることができます。

> [!IMPORTANT]
> `ExcelApiOnline` は、最新の番号付き要件セットのスーパーセットです。

> [!IMPORTANT]
> `ExcelApiOnline 1.1` は、オンライン専用 Api の唯一のバージョンです。 これは、web 上の Excel では、最新バージョンのユーザーが常に1つのバージョンを使用できるためです。

## <a name="recommended-usage"></a>推奨される使用法

`ExcelApiOnline`Api は web 上の Excel でのみサポートされているため、アドインでは、これらの api を呼び出す前に要件セットがサポートされているかどうかを確認する必要があります。 これにより、別のプラットフォームでオンラインのみの API を呼び出すことを回避できます。

```js
if (Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")) {
   // Any API exclusive to the ExcelApiOnline requirement set.
}
```

クロスプラットフォームの要件セットに含まれる API は、チェックを削除または編集する必要があり `isSetSupported` ます。 これにより、他のプラットフォームでアドインの機能が有効になります。 この変更を行うときは、これらのプラットフォームで機能をテストしてください。

> [!IMPORTANT]
> マニフェスト `ExcelApiOnline 1.1` でライセンス認証の要件として指定することはできません。 [Set 要素](../manifest/set.md)で使用する有効な値ではありません。

## <a name="api-list"></a>API リスト

現在、要件セットには Api がありません `ExcelApiOnline` 。 このセットの一部であったすべての Api は、すべてのプラットフォームで使用可能な番号付き要件セットになります。

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel?view=excel-js-online&preserve-view=true)
- [Excel JavaScript プレビュー API](excel-preview-apis.md)
- [Excel JavaScript API の要件セット](excel-api-requirement-sets.md)
