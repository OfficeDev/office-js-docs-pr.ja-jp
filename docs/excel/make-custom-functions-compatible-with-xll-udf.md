---
title: カスタム関数を XLL ユーザー定義関数と互換性があるようにする
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 09914e040c1721dd8b9e91952e5814e7a6b914e5
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356897"
---
# <a name="make-your-custom-functions-compatible-with-xll-user-defined-functions"></a>カスタム関数を XLL ユーザー定義関数と互換性があるようにする

既存の Excel xlls がある場合は、Office アドインで同等のカスタム関数を構築して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。 ただし、Office アドインには、すべての機能が xlls で利用できるわけではありません。 ソリューションで使用されている機能によっては、XLL によって、Excel for Windows の Office アドインカスタム関数よりも優れた操作が提供されることがあります。

同等の xll がユーザーのコンピューターに既にインストールされている場合は、office アドインのカスタム関数ではなく、xll が実行されるように、office アドインを構成することができます。 Excel では、Windows にインストールされているを基にして xll と Office アドインカスタム関数の間で切り替えがシームレスに行われるため、xll は同等と呼ばれます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の xll との互換性を有効にするには、Office アドインのマニフェストで同等の xll を識別します。 その後、Excel での実行時に、Office アドインカスタム関数の代わりに XLL 関数が使用されます。

カスタム関数に対応する xll を設定するには、 `FileName` xll のを指定します。 ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。 ブックは、Windows 上の Excel で開いたときに XLL を使用し、オンラインまたは macOS で開いたときに Office アドインのカスタム関数を使用するようになります。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、この例は完全にコンテキストで指定します。 これらは、 `FileName`それぞれに`ProgID`よって識別されます。 com アドインの互換性の詳細については、「[既存の com アドインと互換性のある Office アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。

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

## <a name="office-add-in-updates"></a>Office アドインの更新プログラム

office アドインに対して同等の XLL を指定すると、Excel は office アドインの更新プログラムの処理を停止します。 ユーザーは、Office アドインの最新の更新プログラムを入手するために XLL をアンインストールする必要があります。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 互換関数のカスタム関数の動作

同じアドインが含まれている xll 関数を含むスプレッドシートが開かれると、xll 関数は、xll 互換のカスタム関数に変換されます。 次の保存時に、これらのファイルは互換モードでファイルに書き込まれ、XLL と Office アドインのカスタム関数 (他のプラットフォーム上の場合) の両方で動作するようになります。

次の表は、xll ユーザー定義関数、xll 互換カスタム関数、Office アドインカスタム関数の機能を比較しています。

|         |XLL のユーザー定義関数 |XLL 互換のカスタム関数 |Office アドインカスタム関数 |
|---------|---------|---------|---------|
| サポートされるプラットフォーム | Windows | Windows、macOS、Excel online | Windows、macOS、Excel online |
| サポートされるファイル形式 | .XLSX、.XLSB、.XLSM、XLS | .XLSX、.XLSB、.XLSM | .XLSX、.XLSB、.XLSM |
| 数式オートコンプリート | いいえ | はい | はい |
| ストリーミング | xlfrtd および XLL コールバックを使用して可能。 | はい | はい |
| 関数のローカライズ | いいえ | いいえ。 名前と ID は、既存の XLL 関数と一致している必要があります。 | はい |
| 揮発性関数 | はい | はい | はい |
| マルチスレッドの再計算のサポート | はい | はい | はい |
| 計算動作 | UI がありません。 計算中に Excel が応答しなくなることがあります。 | ユーザーには #BUSY が表示されます。 を返します。 | ユーザーには #BUSY が表示されます。 を返します。 |
| 要件セット | 該当なし | customfunctions 1.1 のみ | customfunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [既存の COM アドインと互換性のある Office アドインを作成する](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
- [カスタム関数の変更ログ](custom-functions-changelog.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)