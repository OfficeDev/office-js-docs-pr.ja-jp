---
title: Office アドイン用 UX 設計パターン
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 40b36fb138169bdf848e5f58569e6fc3dee8c09b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449010"
---
# <a name="ux-design-patterns-for-office-add-ins"></a>Office アドイン用 UX 設計パターン

Office アドインのユーザー エクスペリエンスの設計では、Office ユーザーにとって魅力的なエクスペリエンスを提供し、既定の Office UI 内でシームレスに適合させることにより Office 全体のエクスペリエンスを拡張する必要があります。  

この UX パターンはコンポーネントで構成されています。 コンポーネントは、お客様がソフトウェアやサービスの要素を操作するのに役立ちます。 ボタン、ナビゲーション、メニューは、整合性のあるスタイルと動作を持つことの多い、一般的なコンポーネントの例です。

Office UI Fabric では、外観も動作も Office の一部のようなコンポーネントを表示します。 Fabric を活用して、Office と簡単に統合します。 アドインに既存のコンポーネント言語がある場合、Fabric のためにその言語を削除する必要はありません。 Office と統合する際に、それを保持する機会を探します。 スタイル要素の入れ替え、競合の削除、ユーザーの混乱を取り除くためのスタイルと動作の採用を行う方法を検討してください。

提供されるパターンは、一般的な顧客シナリオとユーザー エクスペリエンス調査に基づくベスト プラクティス ソリューションです。 それは、アドインの設計と開発のためのクイック エントリ ポイントと、Microsoft とブランド要素のバランスを実現するためのガイダンスの両方を提供することを目的としています。 Microsoft の Fabric 設計言語とパートナー特有のブランドの独自性から得たデザイン要素のバランスを取る、クリーンでモダンなユーザー エクスペリエンスを提供することにより、ユーザーの保持とアドインの採用を向上させることができます。

UX パターン テンプレートを使用して、次のことを行います。

* よくある顧客のシナリオにソリューションとして適用する。
* 設計のベスト プラクティスとして適用する。
* [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) のコンポーネントとスタイルを組み込む。
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
> このドキュメント全体を通して表示されている画面例は、**1366x768**の解像度で設計および表示されています。

## <a name="see-also"></a>関連項目

* [デザインのツールキット](design-toolkits.md)
* [Office UI Fabric](https://developer.microsoft.com/fabric)
* [Office アドイン開発のベスト プラクティス](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [Fabric React の使用の開始](/office/dev/add-ins/design/using-office-ui-fabric-react)
