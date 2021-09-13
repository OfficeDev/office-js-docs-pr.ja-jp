---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等Excel機能を持つ XLL ユーザー定義関数との互換性を有効にする
ms.date: 08/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: 806f920fb6c9a25907fc475cfd29b844ef00f9a8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152992"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>XLL ユーザー定義関数を使用してカスタム関数を拡張する

> [!NOTE]
> XLL アドインは、Excel拡張子 **が .xll** の新しいアドイン ファイルです。 XLL ファイルは、ダイナミック リンク ライブラリ (DLL) ファイルの一種で、このファイルを開くExcel。 XLL アドイン ファイルは、C または C++ で記述する必要があります。 詳細については[、「Excel XLL の開発](/office/client-developer/excel/developing-excel-xlls)」を参照してください。

既存の Excel XLL アドインがある場合は、Excel JavaScript API を使用して同等のカスタム関数アドインをビルドして、Excel on the web や Mac などの他のプラットフォームにソリューション機能を拡張できます。 ただし、Excel JavaScript API アドインには、XLL アドインで使用できるすべての機能が含められません。ソリューションで使用する機能に応じて、XLL アドインは、Excel JavaScript API アドイン (Excel Windows) よりも優れたエクスペリエンスを提供する場合があります。

> [!IMPORTANT]
> COM アドインと XLL ユーザー定義関数 (UDF) の互換性は、Excel (バージョン 1904 以降) Windowsでサポートされます。 COM アドインと XLL ユーザー定義関数 (UDF) の互換性は、アプリケーションまたは Mac ではExcel on the webサポートされていません。

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の XLL アドインとの互換性を有効にするには、JavaScript API アドインのマニフェストで同等の XLL アドインExcel特定します。 Excelで実行すると、Excel JavaScript API アドイン のカスタム関数の代わりに XLL アドインの関数が使用Windows。

カスタム関数に同等の XLL アドインを設定するには `FileName` 、XLL ファイルの値を指定します。 ユーザーが XLL ファイルの関数を含むブックを開くと、Excel互換性のある関数に変換されます。 次に、ブックは Windows の Excel で開かれたときに XLL ファイルを使用し、web または Mac で開かれたときに Excel JavaScript API アドインのカスタム関数を使用します。

次の例は、COM アドインと XLL アドインの両方を、JavaScript API アドイン マニフェスト ファイルの同等物として指定Excel示しています。 多くの場合、両方を指定します。 完全な場合、この例では両方のコンテキストを示します。 これらは、それぞれ、それぞれのユーザー `ProgId` によって `FileName` 識別されます。 要素 `EquivalentAddins` は、終了タグの直前に配置する必要 `VersionOverrides` があります。 COM アドインの互換性の詳細については、「Make your Office アドインを既存の COM アドインと互換性のあるものにする」[を参照してください](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)。

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
> Excel JavaScript API アドインがカスタム関数を XLL アドインと互換性のあるものに宣言すると、後でマニフェストを変更すると、ファイル形式が変更されるので、ユーザーのブックが壊れる可能性があります。

## <a name="custom-function-behavior-for-xll-compatible-functions"></a>XLL 互換関数のカスタム関数の動作

スプレッドシートを開き、同等のアドインが使用可能な場合、アドインの XLL 関数は XLL 互換のカスタム関数に変換されます。 次の保存では、XLL 関数が互換性のあるモードでファイルに書き込まれるので、XLL アドインと Excel JavaScript API アドインの両方のカスタム関数 (他のプラットフォームの場合) で動作します。

次の表は、XLL ユーザー定義関数、XLL 互換カスタム関数、および JavaScript API アドイン Excel機能を比較します。

|         |XLL ユーザー定義関数 |XLL 互換のカスタム関数 |ExcelJavaScript API アドイン のカスタム関数 |
|---------|---------|---------|---------|
| **サポートされるプラットフォーム** | Windows | Windows macOS、Web ブラウザー | Windows macOS、Web ブラウザー |
| **サポートされているファイル形式** | XLSX、XLSB、XLSM、XLS | XLSX、XLSB、XLSM | XLSX、XLSB、XLSM |
| **数式のオートコンプリート** | いいえ | はい | はい |
| **ストリーミング** | xlfRTD コールバックと XLL コールバックを使用して可能です。 | はい | はい |
| **関数のローカライズ** | いいえ | いいえ。 Name と ID は、既存の XLL の関数と一致している必要があります。 | はい |
| **揮発性関数** | はい | はい | はい |
| **マルチスレッド再計算のサポート** | はい | はい | はい |
| **計算動作** | UI なし。 Excelは、計算中に応答しなくなる可能性があります。 | ユーザーには、次の#BUSY! 結果が返されるまで。 | ユーザーには、次の#BUSY! 結果が返されるまで。 |
| **要件セット** | 該当なし | CustomFunctions 1.1 以降 | CustomFunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [Office アドインを既存の COM アドインと互換できるようにする](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
