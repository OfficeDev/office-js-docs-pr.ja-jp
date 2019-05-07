---
title: 既存の COM アドインと互換性のある Excel アドインを作成する
description: Excel アドインと同じ機能を持つ同等の COM アドインとの互換性を有効にする
ms.date: 05/06/2019
localization_priority: Normal
ms.openlocfilehash: 0890e14466a2cd8f5aff2d1bcf307a43cff28127
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628173"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in-preview"></a>既存の COM アドインと互換性のある Office アドインを作成する (プレビュー)

既存の COM アドインがある場合は、Excel アドインで同等の機能を構築して、ソリューション機能をオンラインや macOS などの他のプラットフォームに拡張できます。 ただし、Excel アドインには、COM アドインで使用できるすべての機能が含まれているわけではありません。COM アドインを使用すると、Windows の Excel アドインよりも優れたパフォーマンスを得ることができます。

同等の COM アドインがユーザーのコンピューターに既にインストールされている場合、Office は Excel アドインではなく COM アドインを実行するように、Excel アドインを構成することができます。 COM アドインは、Windows にインストールされているものに応じて、COM アドインと Excel アドインの間でシームレスに移行されるため、"同等" と呼ばれます。

[!include[COM add-in and XLL UDF compatibility requirements note](../includes/xll-compatibility-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>マニフェストで同等の COM アドインを指定する

既存の COM アドインとの互換性を有効にするには、Excel アドインのマニフェストで同等の COM アドインを特定します。 Windows で実行している場合、Office は Excel アドインではなく COM アドインを使用します。

同等の`ProgID` COM アドインのを指定します。 COM アドインがインストールされている場合、Office は、Excel アドインの UI ではなく、COM アドインの UI を使用します。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、この例は完全にコンテキストで指定します。 これらは、 `FileName`それぞれに`ProgID`よって識別されます。 XLL の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。

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

## <a name="equivalent-behavior-for-users"></a>ユーザーの同等の動作

同等の COM アドインが Excel アドインマニフェストで指定されている場合、Office は同等の COM アドインがインストールされている場合、Windows 上で Excel アドインの UI を非表示にします。 これは、オンラインまたは macOS などの他のプラットフォームで Excel アドインの UI に影響を与えることはありません。 Office はリボンボタンを非表示にし、インストールを妨げることはありません。 そのため、Excel アドインは引き続き次の UI の場所に表示されます。

- [ **** アドイン] の下で、技術的にインストールされています。
- リボンマネージャーのエントリとして。

次のシナリオでは、ユーザーが Excel アドインを取得する方法によって実行される処理について説明します。

### <a name="appsource-acquisition-of-an-excel-add-in"></a>Excel アドインの AppSource 取得

ユーザーが AppSource から Excel アドインをダウンロードし、対応する COM アドインが既にインストールされている場合、Office は次のようになります。

1. Excel アドインをインストールします。
2. リボンに Excel アドイン UI を表示しないようにします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="centralized-deployment-of-excel-add-in"></a>Excel アドインの一元展開

管理者が一元展開を使用して Excel アドインをテナントに展開していて、それと同等の COM アドインが既にインストールされている場合、ユーザーは変更を確認する前に Office を再起動する必要があります。 Office を再起動すると、次のようになります。

1. Excel アドインをインストールします。
2. リボンに Excel アドイン UI を表示しないようにします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="document-shared-with-embedded-excel-add-in"></a>埋め込まれた Excel アドインで共有されたドキュメント

ユーザーが COM アドインをインストールしている場合に、埋め込まれた Excel アドインを含む共有ドキュメントを取得すると、Office は次のようになります。

1. Excel アドインを信頼するかどうかをユーザーに確認します。
2. 信頼できる場合は、Excel アドインがインストールされます。
3. リボンに Excel アドイン UI を表示しないようにします。

## <a name="other-com-add-in-behavior"></a>その他の COM アドインの動作

ユーザーが COM アドインをアンインストールすると、Office は、インストールされている excel アドインに対応する Excel アドインの UI を Windows 上に復元します。

Excel アドインに対して同等の COM アドインを指定すると、Office は Excel アドインの更新プログラムの処理を停止します。 ユーザーは、Excel アドインの最新の更新プログラムを取得するために、COM アドインをアンインストールする必要があります。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性を持つようにする](../excel/make-custom-functions-compatible-with-xll-udf.md)
