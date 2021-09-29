---
title: マニフェスト ファイルの EquivalentAddins 要素
description: 同等の COM アドイン、XLL、または両方との下位互換性を指定します。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: d32f67f49d334a75433aec2d079b45a44a04121a
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990811"
---
# <a name="equivalentaddins-element"></a>EquivalentAddins 要素

同等の COM アドイン、XLL、または両方との下位互換性を指定します。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**アドインの種類:** 作業ウィンドウ、メール、カスタム関数

## <a name="syntax"></a>構文

```XML
<EquivalentAddins>
...  
</EquivalentAddins>  
```

## <a name="contained-in"></a>含まれる場所

[VersionOverrides](versionoverrides.md)

## <a name="must-contain"></a>含める必要があるもの

[EquivalentAddin](equivalentaddin.md)

## <a name="see-also"></a>関連項目

- [XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [Office アドインを既存の COM アドインと互換できるようにする](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)