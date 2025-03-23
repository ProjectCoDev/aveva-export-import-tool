# AVEVA Export & Import XML Tool

Un outil graphique moderne pour exporter et importer des objets graphiques au format XML depuis AVEVA System Platform (2023 R2).

![Logo AVEVA](aveva.png)

---

## 🚀 Fonctionnalités

- Interface en Python avec [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- Export d'objets graphiques, templates ou instances en XML ou .aaPKG
- Import d'un graphique XML dans une Galaxy AVEVA
- Mode batch (via fichier `.txt`)
- Affichage du statut + barre de progression
- Design moderne avec thème clair, logo, et icônes

---

## 🖥️ Prérequis

- Windows avec AVEVA System Platform 2023 R2
- Python 3.8+
- Paquets Python :
  - `customtkinter`
  - `Pillow`
  - `pythonnet`

---

## 📦 Installation

Clonez ou téléchargez ce dépôt GitHub :

```bash
https://github.com/<votre-utilisateur>/aveva-export-import-tool.git
```

Puis installez les dépendances :

```bash
pip install -r requirements.txt
```

Assurez-vous que les DLL AVEVA nécessaires sont dans le dossier `dll/`

---

## ⚙️ Utilisation

Lancez simplement :

```bash
python AvevaTool.py
```

- Sélectionnez une Galaxy
- Connectez-vous avec vos identifiants AVEVA
- Choisissez l'objet à exporter ou un fichier `.txt`
- Cliquez sur "Exporter" ou "Importer"

---

## 📜 Licence

Distribué sous licence **MIT**. Voir [LICENSE](LICENSE) pour plus d'infos.

> AVERTISSEMENT : Ce logiciel est fourni "en l'état", sans garantie. Non affilié officiellement à AVEVA.

---

## ✨ À propos

Développé par **Project Collaborative Development EIRL**

> Pour toute contribution, suggestion ou bug, ouvrez une "issue" ou une "pull request". 🤝
