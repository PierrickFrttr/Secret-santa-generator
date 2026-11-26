# 🎅 Secret Santa

Générateur de Secret Santa, ambiance Noël, avec espace organisateur et liens courts.

- **Site :** https://pierrickfrttr.github.io/Secret-santa-generator/
- **Espace organisateur :** https://pierrickfrttr.github.io/Secret-santa-generator/admin
- **Lien d'un tirage :** https://pierrickfrttr.github.io/Secret-santa-generator/noel-cpe

## Fonctionnement
1. L'organisateur se connecte à `/admin` (mot de passe choisi à la première visite), crée un tirage et partage le lien court.
2. Chaque participant touche son prénom, confirme, et ouvre son paquet.
3. L'organisateur suit qui a découvert son prénom (jamais qui offre à qui), peut réinitialiser un participant ou supprimer un tirage.

## Deux types de liens
- **Lien de groupe** : un seul lien dans la conversation, chacun touche son prénom (une seule fois par prénom).
- **Lien par personne** (anti-triche) : chacun reçoit en privé un lien du type `/noel-cpe/lea-x7k2m`, qui ne fonctionne que pour lui.

Le bouton **Vérifier le tirage** de l'espace organisateur contrôle, côté serveur, que chacun offre et reçoit exactement un cadeau, sans révéler qui offre à qui.

## Règles garanties par le serveur
- Personne ne se tire soi-même, chaque prénom n'est tiré qu'une fois (un seul grand cercle, algorithme de Sattolo).
- Un prénom ne peut être utilisé qu'une seule fois, quel que soit l'appareil.
- Le résultat n'est envoyé qu'à la personne qui découvre son prénom.

## Technique
- Pages statiques sur GitHub Pages (`index.html`, `404.html` pour les liens courts).
- Serveur : Google Apps Script (`apps-script/Code.gs`), données dans une feuille Google « Secret Santa – base de données » créée automatiquement.
