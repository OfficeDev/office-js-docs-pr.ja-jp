---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: 3a4793053950fccca74de4b9ebf8998a7d635d67
ms.sourcegitcommit: 65c15a9040279901ea7ff7f522d86c8fddb98e14
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2020
ms.locfileid: "46672688"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>XLL ユーザー定義関数を使用してカスタム関数を拡張する

既存の Excel XLLs を使用している場合は、Excel アドインで同等のカスタム関数を構築して、ソリューション機能をオンラインまたは Mac などの他のプラットフォームに拡張できます。 ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。 ソリューションで使用されている機能によっては、XLL の方が excel の excel アドインカスタム関数よりも優れた操作を提供することがあります。

> [!NOTE]
> COM アドインと XLL の UDF の互換性は、Microsoft 365 サブスクリプションに接続する際に、次のプラットフォームでサポートされています。
> - Excel on the web
> - Windows 版 Excel (バージョン1904以降)
> - Excel on Mac (バージョン13.329 以降)
>
> Web 上の Excel で COM アドインと XLL UDF との互換性を使用するには、Microsoft 365 サブスクリプションまたは [microsoft アカウント](https://account.microsoft.com/account)のいずれかを使用してログインします。 Microsoft 365 サブスクリプションをまだお持ちでない場合は、 [microsoft 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加して、無料の90日更新プログラムの microsoft 365 サブスクリプションをご利用になることができます。

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。 Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。

カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。 ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。 ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは Mac で開いたときに Excel アドインのカスタム関数を使用します。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 通常は、両方を指定します。 完全には、この例の両方がコンテキスト内に表示されます。 これらは、それぞれによって識別され `ProgId` `FileName` ます。 要素は、 `EquivalentAddins` 終了タグの直前に配置する必要があり `VersionOverrides` ます。 COM アドインの互換性の詳細については、「 [既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。

```xml
<VersionOverrides>
  ...
  <EquivalentAddins>
    <EquivalentAddin>
      <ProgId>ContosoCOMAddin</ProgId>
      <Type>COM</Type>
    </EquivalentAddin>

    <EquivalentAddin>
      <FileName>contosofunctions.xll</FileName>
      <Type>XLL</Type>
    </EquivalentAddin>
  </EquivalentAddins>
</VersionOverrides>
```

> [!NOTE]
> アドインでカスタム関数が XLL 互換に宣言されている場合、後でマニフェストを変更すると、ファイル形式が変更されるため、ユーザーのブックが破損する可能性があります。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 互換関数のカスタム関数の動作

アドインの XLL 関数は、スプレッドシートが開かれ、対応するアドインが使用可能な場合に、XLL 互換のカスタム関数に変換されます。 次の保存時に、XLL 関数は互換性のあるモードでファイルに書き込まれます。これにより、他のプラットフォームで使用する場合に、XLL 関数と Excel アドインカスタム関数の両方を使用できるようになります。

次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および Excel アドインカスタム関数の機能を比較しています。

|         |XLL のユーザー定義関数 |XLL 互換のカスタム関数 |Excel アドインのカスタム関数 |
|---------|---------|---------|---------|
| **サポートされるプラットフォーム** | Windows | Windows、macOS、web ブラウザー | Windows、macOS、web ブラウザー |
| **サポートされるファイル形式** | .XLSX、.XLSB、.XLSM、XLS | .XLSX、.XLSB、.XLSM | .XLSX、.XLSB、.XLSM |
| **数式オートコンプリート** | いいえ | はい | はい |
| **ストリーミング** | XlfRTD および XLL コールバックを使用して可能。 | いいえ | はい |
| **関数のローカライズ** | いいえ | いいえ。 名前と ID は、既存の XLL 関数と一致している必要があります。 | はい |
| **揮発性関数** | はい | はい | はい |
| **マルチスレッドの再計算のサポート** | はい | はい | はい |
| **計算動作** | UI がありません。 計算中に Excel が応答しなくなることがあります。 | ユーザーには #BUSY が表示されます。 を返します。 | ユーザーには #BUSY が表示されます。 を返します。 |
| **要件セット** | 該当なし | CustomFunctions 1.1 以降 | CustomFunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [既存の COM アドインと互換性のある Excel アドインを作成する](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
