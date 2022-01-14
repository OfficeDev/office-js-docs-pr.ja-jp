---
title: マニフェスト ファイルの EquivalentAddins 要素
description: 同等の COM アドイン、XLL、または両方との下位互換性を指定します。
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 48f3ef86f71ad3d4f0c759df4583af4cd95e5c5a
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042155"
---
# <a name="equivalentaddins-element"></a>EquivalentAddins 要素

同等の COM アドイン、XLL、または両方との下位互換性を指定します。

[!INCLUDE [Support note for equivalent add-ins feature](../../includes/equivalent-add-in-support-note.md)]

**アドインの種類:** 作業ウィンドウ、メール、カスタム関数

**次の VersionOverrides スキーマでのみ有効です**。

- 作業ウィンドウ 1.0
- メール 1.1

詳細については、「マニフェストの [バージョンオーバーライド」を参照してください](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)。

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