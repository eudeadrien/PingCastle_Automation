# 🔒 AD Security Monitoring – PingCastle Automation

![PowerShell](https://img.shields.io/badge/PowerShell-blue?logo=powershell&logoColor=white)
![MIT License](https://img.shields.io/badge/License-MIT-green)
![Release](https://img.shields.io/badge/Release-v1.0.0-blue)

## 📌 Description
Ce projet automatise l’audit et le suivi de la sécurité Active Directory avec **PingCastle**.  
Il exécute régulièrement des analyses, génère des rapports, crée des graphiques d’évolution et envoie un compte-rendu complet par e-mail.

## 🚀 Fonctionnalités
- ⚡ Exécution planifiée des audits PingCastle  
- 📂 Archivage automatique des rapports (HTML, XML, logs)  
- 📊 Génération de graphiques d’évolution des scores de sécurité  
- ✉️ Envoi d’un reporting professionnel par e-mail (résumé + graphiques + rapport complet en pièce jointe)  
- 🔍 Comparaison automatique avec le rapport précédent pour mettre en avant les changements  

## 🛠️ Technologies utilisées
- **PowerShell** – automatisation & reporting  
- **PingCastle** – outil d’audit AD  
- **System.Drawing** – génération de graphiques  
- **SMTP** – envoi des e-mails automatisés  

## 📦 Structure du projet
```
/PingCastleAutomation
│── /PingCastle # Binaire PingCastle + répertoires Reports & Logs
│ ├── /Logs # Journaux d’exécution
│ └── /Reports # Rapports générés (HTML, XML)
│── 01_RunPingCastle.ps1 # Script principal d’exécution de l'analyse & reporting
│── 02_WeeklyReport.ps1 # Script de reporting hebdomadaire
```

## ⚙️ Utilisation
1. Cloner le repository  
2. Configurer les paramètres SMTP et chemins dans les deux scripts .ps1
3. Planifier les scripts via le Planificateur de tâches Windows  
4. Tout est prêt ! ✅

## 📸 Exemple de résultat
*(contenu à venir)*

## 📜 Licence
Projet open-source sous licence **MIT**
