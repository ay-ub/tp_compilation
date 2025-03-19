# Projet de Compilation 2024/2025

## Analyse et Évaluation d’Expressions Excel en utilisant LEX et YACC

### Description

Ce projet vise à analyser et évaluer des expressions Excel en utilisant **PLY (Python Lex-Yacc)**. Il permet de :

- Effectuer une **analyse lexicale et syntaxique** des expressions Excel.
- Évaluer des expressions avec des **opérations arithmétiques** et des **fonctions Excel**.
- Gérer des **références et plages de cellules** à partir d’un fichier Excel.

---

## 📌 Fonctionnalités

### 🔹 Partie A : Analyse Lexicale et Syntaxique

- Reconnaissance des **références de cellules** (ex: `A1, B2, AA10`)
- Support des **opérateurs arithmétiques** (`+ , - , * , / , ^`)
- Prise en charge des **fonctions Excel** (`SUM, AVERAGE, COUNT, MAX, MIN, etc.`)
- Gestion des **plages de cellules** (`A1:A10`)
- Détection et signalement des erreurs lexicales et syntaxiques

### 🔹 Partie B : Évaluation Sémantique

- Évaluation des **expressions arithmétiques**
- Gestion des **fonctions Excel avec imbrication**
- Support des **plages de cellules** dans les fonctions
- Respect des **priorités des opérations**

---

## 🛠 Installation et Exécution

### 📌 Prérequis

Assurez-vous d’avoir installé :

- **Python 3.x**
- **ply** (Lex-Yacc pour Python)
- **openpyxl** (pour lire les fichiers Excel)
- Un fichier Excel contenant des valeurs (`data.xlsx`)

### 📥 Clonage du Dépôt

```bash
git clone https://github.com/ay-ub/tp_compilation.git
cd tp_compilation
```

### 📦 Installation des Dépendances

```bash
pip install ply openpyxl
```

### ▶️ Exécution du Programme

```bash
python parser.py
```

---

## 📌 Exemples de test

- `SUM(A1:A5) + MAX(B1, AVERAGE(C1:C3, D1))`
- `COUNT(A1:A10) * SUM(B1, B2, MIN(C1:C5))`
- `YEAR("2023-10-05")`

---

## 📌 Auteurs

- **Hadj Youcef Ayyoub** (M1 IL - Département Informatique)

📅 **Année Universitaire 2024-2025**
