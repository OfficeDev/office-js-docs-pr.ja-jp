---
title: XLL ユーザー定義関数を使用してカスタム関数を拡張する
description: カスタム関数と同等の機能を持つ Excel XLL ユーザー定義関数との互換性を有効にする
ms.date: 07/31/2019
localization_priority: Normal
ms.openlocfilehash: 7ec853e5b4d03267e1c9d33d2df8a79d86860095
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235303"
---
# <a name="extend-custom-functions-with-xll-user-defined-functions"></a>XLL ユーザー定義関数を使用してカスタム関数を拡張する

既存の Excel XLLs がある場合は、Excel アドインで同等のカスタム関数を作成して、online や macOS などの他のプラットフォームにソリューション機能を拡張することができます。 ただし、Excel アドインには、xll で利用可能なすべての機能が含まれているわけではありません。 ソリューションで使用されている機能によっては、XLL の方が excel の excel アドインカスタム関数よりも優れた操作を提供することがあります。

> [!NOTE]
> COM アドインと XLL の UDF の互換性は、Office 365 サブスクリプションに接続している場合、次のプラットフォームでサポートされています。
> - Excel on the web
> - Windows 版 Excel (バージョン1904以降)
> - Excel on Mac (バージョン13.329 以降)
> 
> Web 上の Excel で COM アドインと XLL UDF との互換性を使用するには、Office 365 サブスクリプションまたは[Microsoft アカウント](https://account.microsoft.com/account)のいずれかを使用してログインします。 Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 Developer Program](https://developer.microsoft.com/office/dev-program) に参加することで入手できます。

## <a name="specify-equivalent-xll-in-the-manifest"></a>マニフェストで同等の XLL を指定する

既存の XLL との互換性を有効にするには、Excel アドインのマニフェストで同等の XLL を識別します。 Excel では、Windows での実行時に Excel アドインカスタム関数の代わりに XLL 関数が使用されます。

カスタム関数に対応する XLL を設定するには、 `FileName` xll のを指定します。 ユーザーが XLL から関数を含むブックを開くと、Excel は関数を互換性のある関数に変換します。 ブックは、Windows の Excel で開いたときに XLL を使用し、オンラインまたは macOS を開いたときに Excel アドインのカスタム関数を使用します。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、この例は完全にコンテキストで指定します。 これらは、 `FileName`それぞれに`ProgId`よって識別されます。 要素`EquivalentAddins`は、終了`VersionOverrides`タグの直前に配置する必要があります。 COM アドインの互換性の詳細については、「[既存の com アドインと互換性のある Excel アドインを作成](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)する」を参照してください。

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
  <EquivalentAddins>
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
| サポートされるプラットフォーム | Windows | Windows、macOS、Excel on the web | Windows、macOS、Excel on the web |
| サポートされるファイル形式 | .XLSX、.XLSB、.XLSM、XLS | .XLSX、.XLSB、.XLSM | .XLSX、.XLSB、.XLSM |
| 数式オートコンプリート | いいえ | はい | はい |
| ストリーミング | XlfRTD および XLL コールバックを使用して可能。 | いいえ | はい |
| 関数のローカライズ | いいえ | いいえ。 名前と ID は、既存の XLL 関数と一致している必要があります。 | はい |
| 揮発性関数 | はい | はい | はい |
| マルチスレッドの再計算のサポート | はい | はい | はい |
| 計算動作 | UI がありません。 計算中に Excel が応答しなくなることがあります。 | ユーザーには #BUSY が表示されます。 を返します。 | ユーザーには #BUSY が表示されます。 を返します。 |
| 要件セット | なし。 | CustomFunctions 1.1 以降 | CustomFunctions 1.1 以降 |

## <a name="see-also"></a>関連項目

- [既存の COM アドインと互換性のある Excel アドインを作成する](../develop/make-office-add-in-compatible-with-existing-com-add-in.md)
- [チュートリアル: Excel でカスタム関数を作成します。](../tutorials/excel-tutorial-create-custom-functions.md)
