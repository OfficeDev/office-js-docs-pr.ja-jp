---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等Excel機能を持つ XLL ユーザー定義関数との互換性を有効にする
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: b7a2330f7a875c894f371138034314ae99bb0e9393a45c6e8572a97a084fe94e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089315"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>XLL ユーザー定義関数を使用してカスタム関数を拡張する

既存の Excel XLL がある場合は、Excel アドインで同等のカスタム関数を構築して、ソリューション機能をオンラインや Mac などの他のプラットフォームに拡張できます。 ただし、Excelアドインには、すべての機能が XLL で使用できるわけではありません。 ソリューションで使用する機能によっては、XLL は Excel アドインのカスタム関数 (Excel on Windows) よりも優れたエクスペリエンスを提供する場合があります。

> [!NOTE]
> COM アドインと XLL UDF の互換性は、サブスクリプションに接続されている場合、次のプラットフォームMicrosoft 365されます。
>
> - Excel on the web
> - Excel (Windows 1904 以降)
> - Excel (バージョン 13.329 以降)
>
> COM アドインと XLL UDF の互換性を Excel on the web内で使用するには、Microsoft 365 サブスクリプションまたは Microsoft アカウントを使用して[ログインします](https://account.microsoft.com/account)。 Microsoft 365 サブスクリプションをまだ持ってない場合は、Microsoft 365 開発者プログラムに参加することで、90 日間の無料のMicrosoft 365[サブスクリプションを利用できます](https://developer.microsoft.com/office/dev-program)。

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の XLL との互換性を有効にするには、既存のアドインのマニフェストで同等の XLL をExcelします。 Excelで実行すると、アドインのカスタム関数ではなく XLL のExcelを使用Windows。

カスタム関数に同等の XLL を設定するには `FileName` 、XLL の値を指定します。 ユーザーが XLL の関数を含むブックを開くと、Excel互換性のある関数に変換されます。 次に、ブックは Windows の Excel で開かれたときに XLL を使用し、オンラインまたは Mac で開いた場合は Excel アドインのカスタム関数を使用します。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、両方を指定します。 完全な場合、この例では両方のコンテキストを示します。 これらは、それぞれ、それぞれのユーザー `ProgId` によって `FileName` 識別されます。 要素 `EquivalentAddins` は、終了タグの直前に配置する必要 `VersionOverrides` があります。 COM アドインの互換性の詳細については、「Make your Office アドインを既存の COM アドインと互換性のあるものにする」[を参照してください](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)。

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
> アドインがカスタム関数を XLL 互換として宣言すると、後でマニフェストを変更すると、ファイル形式が変更されるので、ユーザーのブックが壊れる可能性があります。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 互換関数のカスタム関数の動作

スプレッドシートを開き、同等のアドインが使用可能な場合、アドインの XLL 関数は XLL 互換のカスタム関数に変換されます。 次の保存では、XLL 関数は互換性のあるモードでファイルに書き込まれるので、XLL と Excel アドインの両方のカスタム関数 (他のプラットフォームの場合) で動作します。

次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、およびアドイン カスタム関数Excel機能を比較します。

|         |XLL ユーザー定義関数 |XLL 互換のカスタム関数 |Excelアドインのカスタム関数 |
|---------|---------|---------|---------|
| **サポートされるプラットフォーム** | Windows | Windows macOS、Web ブラウザー | Windows macOS、Web ブラウザー |
| **サポートされているファイル形式** | XLSX、XLSB、XLSM、XLS | XLSX、XLSB、XLSM | XLSX、XLSB、XLSM |
| **数式のオートコンプリート** | いいえ | はい | 必要 |
| **ストリーミング** | xlfRTD コールバックと XLL コールバックを使用して可能です。 | はい | 必要 |
| **関数のローカライズ** | いいえ | ちがいます。 Name と ID は、既存の XLL の関数と一致している必要があります。 | 必要 |
| **揮発性関数** | はい | はい | 必要 |
| **マルチスレッド再計算のサポート** | はい | はい | 必要 |
| **計算動作** | UI なし。 Excelは、計算中に応答しなくなる可能性があります。 | ユーザーには、次の#BUSY! 結果が返されるまで。 | ユーザーには、次の#BUSY! 結果が返されるまで。 |
| **要件セット** | 該当なし | CustomFunctions 1.1 以降 | CustomFunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [Office アドインを既存の COM アドインと互換できるようにする](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
