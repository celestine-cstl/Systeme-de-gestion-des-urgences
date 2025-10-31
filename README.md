# Système de gestion des urgences
**Conception d’un flux urgences réaliste sous Excel/VBA : admission, tri infirmier, prise en charge médicale, gestion du personnel, archivage et statistiques automatisées.**

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-00A300?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![Word](https://img.shields.io/badge/Word-2B579A?style=for-the-badge\&logo=microsoft-word\&logoColor=white)

> L’application est réalisée **entièrement sous Excel/VBA** (formulaires, règles métier, validations, priorisation, archivage et statistiques), **Word** est utilisé pour la rédaction et la synthèse des résultats.

<br>

## 🛠️ Compétences mobilisées
- **Modélisation de processus hospitaliers** : parcours patient réaliste (accueil, tri infirmier, prise en charge médicale, sortie/archivage)
- **Conception d’interface utilisateur VBA** : formulaires guidant l’agent (Home, Admission, Priorisation, Prise en charge, Statistiques)
- **Data integrity & contrôle métier** : unicité identité patient, générateur d’IEP (identifiant d’épisode), anticolision arrivée/départ du personnel, anti-doublons patients
- **Algorithmes de priorisation** : fonctions CalculerPriorite et CalculerPrioriteMedicale
- **Gestion RH temps réel** : présence/absence, début/fin de service
- **Reporting & analytics** : module statistiques hebdo/mensuel (regroupements par médecin/infirmier/mode d’arrivée/pôle/date)
- **Communication écrite** : synthèses claires et structurées dans Word.

<br>

## 📑 Contenu du projet 

**Accueil & Admission :** Enregistrement des données d’identité et d’arrivée, création d’un IEP par passage, prévention des doublons. La page "Home" oriente l’utilisateur étape par étape.

**Priorisation infirmière :** Évaluation clinique avec paramètres vitaux, détermination/ajustement de la priorité et orientation vers le pôle de soins. Sélection automatique du patient prioritaire via CalculerPriorite et réévaluation dans le temps. Gestion du retour à domicile.

**Prise en charge médicale :** Formulaire médecin initialisant listes selon le pôle ; priorité aux arrivées SMUR, puis patients non-SMUR via CalculerPrioriteMedicale. Saisie de la pathologie, décision de sortie et traçabilité complète.

**Gestion du personnel :** Tableau de bord présences, ajout d’un praticien et suppression d’un membre sortant via formulaires dédiés.

**Archivage & traçabilité :** Tous les séjours sortis sont archivés automatiquement, alimentant le module d’analyse.

**Statistiques automatisées :** Génération des tableaux/graphes
