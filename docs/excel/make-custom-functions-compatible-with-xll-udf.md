---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする (プレビュー)
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: 3e1782c5df227d3e173f4291ba88f2057200b1c5
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33951887"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions-preview"></a>XLL ユーザー定義関数を使用してカスタム関数を拡張する (プレビュー)

既存の Excel XLLs がある場合は、Excel アドインで同等のカスタム関数を作成して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。 ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。 ソリューションで使用されている機能によっては、XLL の方が excel の excel アドインカスタム関数よりも優れた操作を提供することがあります。

[!include[COM add-in and XLL UDF compatibility note](../includes/xll-compatibility-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。 Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。

カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。 ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。 ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは macOS を開いたときに Excel アドインのカスタム関数を使用します。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、この例は完全にコンテキストで指定します。 これらは、 `FileName`それぞれに`ProgID`よって識別されます。 COM アドインの互換性の詳細については、「[既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。

```xml
<VersionOverrides>
...
<EquivalentAddins>
  <EquivalentAddin>
    <ProgID>ContosoCOMAddin</ProgID>
    <Type>COM</Type>
  </EquivalentAddin>

  <EquivalentAddin>
    <FileName>contosofunctions.xll</FileName>
    <Type>XLL</Type>
  </EquivalentAddin>
<EquivalentAddins>
...
</VersionOverrides>
```

> [!NOTE]
> アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。

## <a name="excel-add-in-updates"></a>Excel アドインの更新プログラム

Excel アドインに対して同等の XLL を指定すると、excel アドインの更新プログラムの処理は中止されます。 ユーザーは、Excel アドインの最新の更新プログラムを取得するために、XLL をアンインストールする必要があります。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 互換関数のカスタム関数の動作

同じアドインが含まれている XLL 関数を含むスプレッドシートが開かれると、xll 関数は、XLL 互換のカスタム関数に変換されます。 次の保存時に、これらのファイルは互換モードでファイルに書き込まれます。これにより、(他のプラットフォームでの場合) XLL と Excel アドインの両方のカスタム機能を使用できるようになります。

次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および Excel アドインカスタム関数の機能を比較しています。

|         |XLL のユーザー定義関数 |XLL 互換のカスタム関数 |Excel アドインのカスタム関数 |
|---------|---------|---------|---------|
| サポートされるプラットフォーム | Windows | Windows、macOS、Excel online | Windows、macOS、Excel online |
| サポートされるファイル形式 | .XLSX、.XLSB、.XLSM、XLS | .XLSX、.XLSB、.XLSM | .XLSX、.XLSB、.XLSM |
| 数式オートコンプリート | いいえ | はい | はい |
| ストリーミング | XlfRTD および XLL コールバックを使用して可能。 | いいえ | はい |
| 関数のローカライズ | 不要 | いいえ。 名前と ID は、既存の XLL 関数と一致している必要があります。 | はい |
| 揮発性関数 | はい | はい | はい |
| マルチスレッドの再計算のサポート | はい | はい | はい |
| 計算動作 | UI がありません。 計算中に Excel が応答しなくなることがあります。 | ユーザーには #BUSY が表示されます。 を返します。 | ユーザーには #BUSY が表示されます。 を返します。 |
| 要件セット | N/A | CustomFunctions 1.1 以降 | CustomFunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [既存の COM アドインと互換性のある Excel アドインを作成する](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
