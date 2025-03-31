# 🖥️ System Information Report Generator - VBScript

**System Information Report Generator** est un outil puissant écrit en VBScript qui génère un rapport détaillé et complet des informations système de votre ordinateur Windows. Obtenez instantanément un aperçu complet de votre configuration matérielle, logicielle et réseau en un seul clic.

![Banner](https://o2cloud.fr/logo/o2Cloud.png)

## ✨ Fonctionnalités

- 💻 **Analyse matérielle complète** - CPU, RAM, cartes graphiques, disques, périphériques, etc.
- 🔍 **Informations détaillées sur le système d'exploitation** - version, build, état d'activation
- 👥 **Liste des utilisateurs** - comptes locaux et domaine avec leurs privilèges
- ⚙️ **Inventaire des processus** - processus en cours d'exécution avec leur consommation de ressources
- 🌐 **Configuration réseau** - adaptateurs, adresses IP, DNS, passerelles
- 📊 **Génération automatique de rapports** - formats TXT ou HTML pour un partage facile
- 🛠️ **Aucune installation requise** - script autonome prêt à l'emploi
- 🔧 **Hautement personnalisable** - ajustez le script selon vos besoins spécifiques

## 📋 Pré-requis

- Système d'exploitation Windows (Windows 7 ou plus récent)
- Windows Script Host activé (activé par défaut sur la plupart des systèmes Windows)

## 🚀 Utilisation

1. Téléchargez ou clonez ce dépôt :
   ```bash
   git clone https://github.com/o2Cloud-fr/System-Information-Report-Generator.git
   ```

2. Exécutez le script en double-cliquant sur le fichier `SystemInfoGenerator.vbs` ou via la ligne de commande :
   ```bash 
   cscript SystemInfoGenerator.vbs
   ```

3. Le rapport sera généré dans le même répertoire que le script.

## 📚 Documentation

Le script utilise WMI (Windows Management Instrumentation) pour collecter des informations système détaillées, notamment :

- Informations sur le processeur (fabricant, modèle, fréquence)
- Configuration de la mémoire RAM (capacité totale, modules installés)
- Détails des disques (capacité, espace libre, partitions)
- Informations sur le système d'exploitation (version, build, activation)
- Configuration réseau (adaptateurs, adresses IP, DNS)
- Liste des programmes installés
- Services en cours d'exécution et leur état

## 👨‍💻 Auteurs

- [@MyAlien](https://www.github.com/MyAlien)
- [@o2Cloud](https://www.github.com/o2Cloud-fr)

## 🔖 Badges

[![Apache License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://github.com/o2Cloud-fr/System-Information-Report-Generator/blob/main/LICENSE)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)](https://github.com/o2Cloud-fr/System-Information-Report-Generator)
[![VBScript](https://img.shields.io/badge/Language-VBScript-yellow.svg)](https://github.com/o2Cloud-fr/System-Information-Report-Generator)
[![o2Cloud](https://img.shields.io/badge/Powered%20by-o2Cloud-orange.svg)](https://o2cloud.fr/)

## 🤝 Contribution

Les contributions sont toujours les bienvenues !

Consultez le fichier `contributing.md` pour découvrir comment contribuer à ce projet.
Veuillez respecter le `code of conduct` du projet.

## 💬 Feedback

Si vous avez des commentaires ou des suggestions, n'hésitez pas à nous contacter à github@o2cloud.fr

## 🔗 Liens

[![portfolio](https://img.shields.io/badge/my_portfolio-000?style=for-the-badge&logo=ko-fi&logoColor=white)](https://vcard.o2cloud.fr/)
[![linkedin](https://img.shields.io/badge/linkedin-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/remi-simier-2b30142a1/)

## 🛠️ Compétences

- VBScript
- WMI (Windows Management Instrumentation)
- Scripting Windows

## 📝 Licence

[Apache-2.0 license](https://github.com/o2Cloud-fr/System-Information-Report-Generator/blob/main/LICENSE)

## 🔄 Projets connexes

Voici quelques projets similaires ou complémentaires :
- [GitHub o2Cloud](https://github.com/o2Cloud-fr?tab=repositories)
- [Awesome README](https://github.com/o2Cloud-fr/System-Information-Report-Generator/blob/main/README.md)

## 🗺️ Feuille de route

- Ajouter un mode d'exportation CSV pour l'analyse de données
- Support pour l'exportation au format JSON
- Intégration de la détection des vulnérabilités de sécurité
- Possibilité de générer des rapports différentiels (comparaison entre deux scans)
- Interface graphique simplifiée pour la visualisation des rapports

## 🆘 Support

Pour obtenir de l'aide, envoyez un e-mail à github@o2cloud.fr ou rejoignez notre canal Slack.

## 💼 Utilisé par

Ce projet est utilisé par les entreprises suivantes :
- o2Cloud
- MyAlienTech
