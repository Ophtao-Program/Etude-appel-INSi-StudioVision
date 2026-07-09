# Étude INSi — rapport statistique

```
========================================================================
  ÉTUDE INSi — RAPPORT STATISTIQUE
========================================================================
  Fichier ................ C:\INFO-BOX-CV\Etude-appel-INSi-StudioVision\Etude-appel-INSi-StudioVision-main\etude_insi_resultats.json
  Enregistrements ....... 3762  (patients distincts : 3762)
  Consultations ......... du 2026-01-02 au 2026-07-10
  Exécution ............. du 2026-07-08 00:32:36 au 2026-07-09 08:49:29

------------------------------------------------------------------------
  1. RÉSULTAT PRINCIPAL — appels INSi ayant reçu une réponse
------------------------------------------------------------------------
  Appels aboutis ........ 3577  (95,1 % des patients analysés)

   00 — patient trouvé                                 2691    75,2 %
      █████████████████████████████·········
   01 — patient NON trouvé (traits d'identité diverge   872    24,4 %
      █████████·····························
   02 — plusieurs patients trouvés                       14     0,4 %
      ······································

  ➜ TAUX DE RÉUSSITE INSi (00 / aboutis) ......... 75,2 %
  ➜ Traits d'identité divergents (01 / aboutis) .. 24,4 %
  ➜ Ambiguïté (02 / aboutis) .................... 0,4 %

------------------------------------------------------------------------
  2. RÉPARTITION COMPLÈTE (tous les cas)
------------------------------------------------------------------------
   00 — patient trouvé                                 2691    71,5 %
      ███████████████████████████···········
   01 — patient NON trouvé (traits d'identité diverge   872    23,2 %
      █████████·····························
   sexe absent, aucun sexe ne fonctionne                 90     2,4 %
      █·····································
   échec d'ouverture de la fiche                         76     2,0 %
      █·····································
   sexe absent, pas de n°SS (non résolu)                 15     0,4 %
      ······································
   02 — plusieurs patients trouvés                       14     0,4 %
      ······································
   aucune réponse du téléservice                          3     0,1 %
      ······································
   sous-formulaire non ouvert (incident Access)           1     0,0 %
      ······································

------------------------------------------------------------------------
  3. FICHES SANS SEXE — résolution automatique
------------------------------------------------------------------------
  Fiches au sexe vide ... 532  (14,1 % des patients analysés)
  Sexe corrigé .......... 427  (80,3 % des fiches sans sexe)
  Restées sans sexe ..... 105

  Par méthode de déduction :
   enfant (essais F/M)                  257   résolus :  257  (100,0 %)
   déduit du n°SS                       252   résolus :  167  ( 66,3 %)
   pas de n°SS (essais F/M)              18   résolus :    3  ( 16,7 %)
   n°SS non exploitable (essais F/M)      5   résolus :    0  (  0,0 %)

  Sexe finalement retenu : F = 225  M = 202

------------------------------------------------------------------------
  4. CROISEMENTS (sur les appels aboutis)
------------------------------------------------------------------------

  Par sexe de la fiche :
                        total       00 trouvé   01 non trouvé    02 plusieurs
    ------------------------------------------------------------------------
    M                    1473    1446 (98,2%)       20 (1,4%)        7 (0,5%)
    F                    2104    1245 (59,2%)     852 (40,5%)        7 (0,3%)

  Par tranche d'âge :
                        total       00 trouvé   01 non trouvé    02 plusieurs
    ------------------------------------------------------------------------
    0–17 ans              646     632 (97,8%)       10 (1,5%)        4 (0,6%)
    18–39 ans             553     501 (90,6%)       49 (8,9%)        3 (0,5%)
    40–59 ans             816     611 (74,9%)     203 (24,9%)        2 (0,2%)
    60–74 ans             774     508 (65,6%)     262 (33,9%)        4 (0,5%)
    75 ans et +           788     439 (55,7%)     348 (44,2%)        1 (0,1%)

  Selon la présence du n°SS :
                        total       00 trouvé   01 non trouvé    02 plusieurs
    ------------------------------------------------------------------------
    n°SS présent         3547    2673 (75,4%)     860 (24,2%)       14 (0,4%)
    n°SS absent            30      18 (60,0%)      12 (40,0%)               –

  Selon le titre :
                        total       00 trouvé   01 non trouvé    02 plusieurs
    ------------------------------------------------------------------------
    Enfant                557     549 (98,6%)        5 (0,9%)        3 (0,5%)
    (autre)              3020    2142 (70,9%)     867 (28,7%)       11 (0,4%)

------------------------------------------------------------------------
  5. IDENTIFIANT NATIONAL DE SANTÉ (INS)
------------------------------------------------------------------------
  INS obtenu ............ 2691  (71,5 % des patients analysés, 100,0 % des 00)

------------------------------------------------------------------------
  6. APPELS AU TÉLÉSERVICE
------------------------------------------------------------------------
  Appels INSi émis ...... 3896
  Patients à 2 appels ... 214  (résolution du sexe)

------------------------------------------------------------------------
  7. QUALITÉ DE L'EXÉCUTION
------------------------------------------------------------------------
  Incidents à reprendre . 80  (2,1 %)
     échec d'ouverture de la fiche                     76
     aucune réponse du téléservice                      3
     sous-formulaire non ouvert (incident Access)       1
     ➜ relancer « 5 - Etude INSi (complete) » : ces cas seront refaits.
  Incohérences nom/fiche  3

========================================================================
```
