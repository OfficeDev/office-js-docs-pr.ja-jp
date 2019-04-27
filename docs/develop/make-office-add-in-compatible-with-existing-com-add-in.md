---
title: 既存の COM アドインと互換性のある Office アドインを作成する
description: Office アドインと同じ機能を持つ同等の COM アドインとの互換性を有効にする
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 8f3780814163cc4dd21311b362d1d821a14b3e80
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356895"
---
# <a name="make-your-office-add-in-compatible-with-an-existing-com-add-in"></a>既存の COM アドインと互換性のある Office アドインを作成する

既存の COM アドインがある場合は、Office アドインで同等の機能を構築して、ソリューション機能を online や macOS などの他のプラットフォームに拡張できます。 ただし、Office アドインには、COM アドインで使用できるすべての機能が含まれているわけではありません。COM アドインでは、Excel、Word、および PowerPoint の Office アドインよりも優れた機能を提供する場合があります。

同等の com アドインがユーザーのコンピューターに既にインストールされている場合は office アドインを構成できます。 office は、office アドインではなく、com アドインを実行します。 com アドインは、office が Windows にインストールされているものに応じて、com アドインと office アドインをシームレスに移行するため、"同等" と呼ばれます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="specify-an-equivalent-com-add-in-in-the-manifest"></a>マニフェストで同等の COM アドインを指定する

既存の com アドインとの互換性を有効にするには、Office アドインのマニフェストで同等の com アドインを特定します。 Windows で実行している場合、office は office アドインではなく、COM アドインを使用します。

同等の`ProgID` COM アドインのを指定します。 これで、com アドインをインストールするときに、office アドインの ui ではなく、com アドインの ui が使用されます。

次の例は、COM アドインと XLL の両方を同等として指定する方法を示しています。 多くの場合、この例は完全にコンテキストで指定します。 これらは、 `FileName`それぞれに`ProgID`よって識別されます。 xll の互換性の詳細については、「 [xll ユーザー定義関数と互換性のあるカスタム関数を作成する](../excel/make-custom-functions-compatible-with-xll-udf.md)」を参照してください。

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

office アドインマニフェストで同等の com アドインが指定されている場合、office は、対応する com アドインがインストールされている場合、Windows 上で office アドインの UI を表示しません。 これは、online や macOS などの他のプラットフォームで Office アドインの UI に影響を与えることはありません。 Office はリボンボタンを非表示にし、インストールを妨げることはありません。 そのため、Office アドインは引き続き次の UI の場所に表示されます。

- [ **** アドイン] の下で、技術的にインストールされています。
- リボンマネージャーのエントリとして。

次のシナリオでは、ユーザーが Office アドインを取得する方法によって実行される処理について説明します。

### <a name="appsource-acquisition-of-an-office-add-in"></a>Office アドインの appsource 取得

ユーザーが appsource から Office アドインをダウンロードし、対応する COM アドインが既にインストールされている場合、office は次のようになります。

1. Office アドインをインストールします。
2. リボンで Office アドイン UI を非表示にします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="centralized-deployment-of-office-add-in"></a>Office アドインの一元展開

管理者が一元展開を使用して office アドインをテナントに展開していて、それと同等の COM アドインが既にインストールされている場合、ユーザーは変更を確認する前に office を再起動する必要があります。 Office を再起動すると、次のようになります。

1. Office アドインをインストールします。
2. リボンで Office アドイン UI を非表示にします。
3. [COM アドイン] リボンボタンをポイントするユーザーの呼び出しを表示します。

### <a name="document-shared-with-embedded-office-add-in"></a>埋め込まれた Office アドインと共有されたドキュメント

ユーザーが COM アドインをインストールしていて、office アドインが埋め込まれた共有ドキュメントを取得した場合、office は次のようになります。

1. Office アドインを信頼するかどうかをユーザーに確認します。
2. 信頼できる場合は、Office アドインがインストールされます。
3. リボンで Office アドイン UI を非表示にします。

## <a name="other-com-add-in-behavior"></a>その他の COM アドインの動作

ユーザーが COM アドインをアンインストールすると、office アドインの UI は、インストールされている office アドインに対応する Windows 上で復元されます。

office アドインに対して同等の COM アドインを指定すると、office アドインの更新プログラムの処理は中止されます。 ユーザーは、Office アドインの最新の更新プログラムを取得するために、COM アドインをアンインストールする必要があります。

## <a name="see-also"></a>関連項目

- [カスタム関数を XLL ユーザー定義関数と互換性を持つようにする](../excel/make-custom-functions-compatible-with-xll-udf.md)
