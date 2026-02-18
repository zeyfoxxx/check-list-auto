# check-list-auto
automatisation de la check-list WAN du matin

Il s'agit d'un code Python qui grâce à différentes librairies (OS, openpyxl, ...) va faire les premières manipulations les plus basiques de manière automatique.

Ils se connectent à Outlook dans les fichiers qui lui sont indiqués, ils recherchent les fichiers en fonction des mot-clé qui leur est indiqué.

Les documents/fichiers qu'il a trouvés, il vient les télécharger dans le fichier temps que je lui fais créer justement.

Dans le fichier par défaut, il y a aussi un template de la check-list (donc un fichier .xlsx) où je viens lui indiquer les cases où il doit rajouter ce qu'il a trouvé.

Attention, il ne rajoute pas dans le fichier template, il crée une copie de celui-ci que je pourrais supprimer à la fin de la check-list, comme ça pas besoin de lui redonner le template à chaque fois.

Grâce à une chaîne de caractères, je viens vérifier dans une case spécifique les faux positifs, si seulement les faux positifs sont présents, ils les remplacent par Ø.

Bien sûr, il vient faire des commentaires à l'avancement de la check-list pour savoir s'il y a des problèmes rencontrés.
