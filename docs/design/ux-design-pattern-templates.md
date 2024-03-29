---
title: Office アドイン用 UX 設計パターン
description: ナビゲーション、認証、初回実行、ブランド化のパターンなど、Officeアドインの UI デザイン パターンの概要を確認します。
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9349096ecb03e4fa7980b9430f3940b5d09bdab2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743115"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Office アドイン用 UX 設計パターン

Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。  

この UX パターンはコンポーネントで構成されています。 コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。 ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。

[Fluent UI](using-office-ui-fabric-react.md) Reactは、JS のフレームワークに依存しないコンポーネントと同様に、Office の一部のように見え、Office UI Fabric[動作します](fabric-core.md)。 いずれかのコンポーネントセットを利用して、複数のコンポーネントと統合Office。 または、アドインに独自の既存のコンポーネント言語がある場合は、その言語を破棄する必要があります。 Office と統合する際に、それを保持する機会を探します。 スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。

提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。 これらは、アドインの設計と開発に関する簡単なエントリ ポイントと、Microsoft ブランド要素と独自のブランド要素のバランスを取るガイダンスの両方を提供することを目的としています。 Microsoft の Fluent UI デザイン言語とパートナー固有のブランド ID との間でデザイン要素のバランスを取るクリーンでモダンなユーザー エクスペリエンスを提供することで、ユーザーの保持とアドインの導入が向上する可能性があります。

UX パターン テンプレートを使用して、次のことを行います。

* よくある顧客のシナリオにソリューションとして適用する。
* 設計のベスト プラクティスとして適用する。
* UI [Fluentスタイル](https://developer.microsoft.com/fluentui#/get-started)を組み込む。
* Office の既定の UI に視覚的に溶け込むアドインをビルドする。
* UX を観念化および可視化する。

## <a name="getting-started"></a>はじめに

パターンは、キーの動作またはアドインに共通のエクスペリエンスによって構成されます。 主なグループは次のとおりです。

* [最初の実行エクスペリエンス (FRE)](../design/first-run-experience-patterns.md)
* [認証](../design/authentication-patterns.md)
* [ナビゲーション](../design/navigation-patterns.md)
* [ブランド デザイン](../design/branding-patterns.md)

各グループを参照して、ベスト プラクティスを使ってアドインを設計する方法を理解します。

> [!NOTE]
> このドキュメント全体を通して表示されている画面例は、**1366x768** の解像度で設計および表示されています。

## <a name="see-also"></a>関連項目

* [デザインのツール キット](design-toolkits.md)
* [Fluent UI](https://developer.microsoft.com/fluentui#)
* [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
* [Office アドインの Fluent UI React](using-office-ui-fabric-react.md)
