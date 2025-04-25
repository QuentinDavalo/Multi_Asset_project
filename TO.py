import pandas as pd
import numpy as np
import statsmodels.api as sm
from math import sqrt
import os

# Constantes pour l'annualisation
JOURS_TRADING_ANNEE = 252

# Définition des actions par secteur
secteurs = {
    "Communication Services": ["PUBLICIS GROUPE"],
    "Consumer Discretionary": ["INDUSTRIA DE DISENO TEXTIL", "MICHELIN (CGDE)"],
    "Consumer Staples": ["BEIERSDORF AG", "HENKEL AG & CO KGAA VOR-PREF"],
    "Financials": ["LEGAL & GENERAL GROUP PLC", "LONDON STOCK EXCHANGE GROUP", "AVIVA PLC", "ADYEN NV"],
    "Health Care": ["STRAUMANN HOLDING AG-REG", "SANOFI"],
    "Industrials": ["KONE OYJ-B", "SIEMENS ENERGY AG", "AIRBUS SE", "RHEINMETALL AG"],
    "Information Technology": ["ASML HOLDING NV"],
    "Materials": ["RIO TINTO PLC"],
    "Utilities": ["RWE AG"]
}

# Mapping des noms complets aux symboles boursiers
symbole_mapping = {
    "PUBLICIS GROUPE": "PUB.PA",
    "INDUSTRIA DE DISENO TEXTIL": "ITX.MC",
    "MICHELIN (CGDE)": "ML.PA",
    "BEIERSDORF AG": "BEI.DE",
    "HENKEL AG & CO KGAA VOR-PREF": "HEN3.DE",
    "LEGAL & GENERAL GROUP PLC": "LGEN.L",
    "LONDON STOCK EXCHANGE GROUP": "LSEG.L",
    "AVIVA PLC": "AV.L",
    "ADYEN NV": "ADYEN.AS",
    "STRAUMANN HOLDING AG-REG": "STMN.SW",
    "SANOFI": "SAN.PA",
    "KONE OYJ-B": "KNEBV.HE",
    "SIEMENS ENERGY AG": "ENR.DE",
    "AIRBUS SE": "AIR.PA",
    "RHEINMETALL AG": "RHM.DE",
    "ASML HOLDING NV": "ASML.AS",
    "RIO TINTO PLC": "RIO.L",
    "RWE AG": "RWE.DE"
}

# Mapping inverse pour retrouver les noms d'entreprise à partir des symboles
nom_mapping = {v: k for k, v in symbole_mapping.items()}


# Fonction pour charger les données Excel
def charger_donnees(fichier_excel):
    try:
        # Charger les prix de clôture et rendements
        prix_quotidiens = pd.read_excel(fichier_excel, sheet_name='Prix Quotidiens', index_col=0)
        rendements_journaliers = pd.read_excel(fichier_excel, sheet_name='Rendements Journaliers', index_col=0)

        return prix_quotidiens, rendements_journaliers
    except Exception as e:
        print(f"Erreur lors du chargement des données: {e}")
        # Vérifier si des fichiers CSV existent (alternative)
        if os.path.exists('prix_quotidiens_5ans.csv') and os.path.exists('rendements_journaliers_5ans.csv'):
            prix_quotidiens = pd.read_csv('prix_quotidiens_5ans.csv', index_col=0)
            rendements_journaliers = pd.read_csv('rendements_journaliers_5ans.csv', index_col=0)
            return prix_quotidiens, rendements_journaliers
        else:
            raise e


# Fonction pour trouver le secteur d'un symbole
def trouver_secteur(symbole):
    nom_entreprise = nom_mapping.get(symbole)
    if not nom_entreprise:
        return "Autre"

    for secteur, entreprises in secteurs.items():
        if nom_entreprise in entreprises:
            return secteur

    return "Autre"


# Fonction pour calculer le rendement géométrique
def calculer_rendement_geometrique(serie_prix):
    if len(serie_prix) <= 1:
        return np.nan

    # Calcul du rendement total sur la période
    rendement_total = serie_prix.iloc[-1] / serie_prix.iloc[0] - 1

    # Calcul du nombre d'années (approximation)
    nb_jours = len(serie_prix)
    nb_annees = nb_jours / JOURS_TRADING_ANNEE

    # Rendement géométrique annualisé
    return (1 + rendement_total) ** (1 / nb_annees) - 1


# Fonction pour calculer les métriques de régression et de volatilité
def calculer_metriques_avancees(rendements_action, rendements_indice, prix_action=None):
    if len(rendements_action) < 30 or len(rendements_indice) < 30:
        return None

    # Nettoyer les données pour éliminer les valeurs manquantes
    data = pd.DataFrame({'y': rendements_action, 'x': rendements_indice}).dropna()
    if len(data) < 30:
        return None

    # Extraire les séries nettoyées
    y = data['y']
    x = data['x']

    # Calculer les statistiques de base
    vol_totale = y.std() * sqrt(JOURS_TRADING_ANNEE)  # Volatilité annualisée de l'action
    vol_indice = x.std() * sqrt(JOURS_TRADING_ANNEE)  # Volatilité annualisée de l'indice
    correlation = y.corr(x)  # Corrélation avec l'indice

    # Régression
    X = sm.add_constant(x)
    model = sm.OLS(y, X).fit()

    # Extraire les paramètres
    alpha = model.params[0] * JOURS_TRADING_ANNEE  # Alpha annualisé
    beta = model.params[1]  # Bêta
    r_squared = model.rsquared  # R²

    # Calcul des volatilités décomposées
    vol_systematique = beta * vol_indice  # Volatilité systématique
    vol_residuelle = sqrt(max(0,
                              vol_totale ** 2 - vol_systematique ** 2))  # Volatilité résiduelle (avec protection contre valeurs négatives)

    # Calcul du rendement géométrique annualisé si les prix sont fournis
    if prix_action is not None and len(prix_action) > 1:
        rendement_geo = calculer_rendement_geometrique(prix_action)
    else:
        rendement_geo = np.nan

    return {
        'Alpha': alpha,  # Format décimal
        'Beta': beta,
        'R-squared': r_squared,
        'Correlation': correlation,
        'Rendement_Geo_Annualisé': rendement_geo,
        'Volatilité_Totale': vol_totale,  # Format décimal
        'Volatilité_Systématique': vol_systematique,  # Format décimal
        'Volatilité_Résiduelle': vol_residuelle  # Format décimal
    }


def main():
    try:
        # Charger les données
        fichier_excel = 'resultats_actions_5ans.xlsx'
        print(f"Chargement des données depuis {fichier_excel}...")
        prix_quotidiens, rendements_journaliers = charger_donnees(fichier_excel)

        # Identifier l'indice de référence (première colonne normalement)
        indice_ref = prix_quotidiens.columns[0]
        print(f"Indice de référence: {indice_ref}")

        # Filtrer les colonnes pour n'inclure que les symboles sélectionnés
        symboles_selectionnés = list(symbole_mapping.values())
        symboles_a_analyser = [indice_ref] + [s for s in symboles_selectionnés if s in rendements_journaliers.columns]

        # Créer un DataFrame pour stocker les résultats par secteur
        colonnes_resultats = ['Secteur', 'Entreprise', 'Symbole', 'Alpha', 'Beta', 'R-squared',
                              'Correlation', 'Rendement_Geo_Annualisé', 'Volatilité_Totale',
                              'Volatilité_Systématique', 'Volatilité_Résiduelle']
        resultats = pd.DataFrame(columns=colonnes_resultats)

        # Analyser l'indice d'abord
        print(f"\nAnalyse de l'indice: {indice_ref}")

        rendements_indice = rendements_journaliers[indice_ref].dropna()
        prix_indice = prix_quotidiens[indice_ref].dropna()
        metriques_indice = calculer_metriques_avancees(rendements_indice, rendements_indice, prix_indice)

        if metriques_indice:
            # Ajouter les résultats pour l'indice
            nouvelle_ligne = pd.DataFrame([{
                'Secteur': "Indice",
                'Entreprise': "Eurostoxx",
                'Symbole': indice_ref,
                'Alpha': 0,  # Par définition, l'alpha de l'indice par rapport à lui-même est 0
                'Beta': 1,  # Par définition, le bêta de l'indice par rapport à lui-même est 1
                'R-squared': 1,  # Par définition, le R² de l'indice par rapport à lui-même est 1
                'Correlation': 1,  # Par définition, la corrélation de l'indice avec lui-même est 1
                'Rendement_Geo_Annualisé': metriques_indice['Rendement_Geo_Annualisé'],
                'Volatilité_Totale': metriques_indice['Volatilité_Totale'],
                'Volatilité_Systématique': metriques_indice['Volatilité_Totale'],
                # Toute la volatilité est systématique
                'Volatilité_Résiduelle': 0  # Pas de volatilité résiduelle
            }])

            resultats = pd.concat([resultats, nouvelle_ligne], ignore_index=True)

        # Analyser chaque action par secteur
        for symbole in symboles_a_analyser:
            if symbole == indice_ref:
                continue  # Déjà traité

            # Trouver le secteur et le nom de l'entreprise
            secteur = trouver_secteur(symbole)
            entreprise = nom_mapping.get(symbole, symbole)

            print(f"Analyse de {entreprise} ({symbole}) - Secteur: {secteur}")

            # Extraire les rendements et les prix
            rendements_action = rendements_journaliers[symbole].dropna()
            rendements_indice_filtres = rendements_journaliers[indice_ref].loc[rendements_action.index]
            prix_action = prix_quotidiens[symbole].dropna()

            # Calculer les métriques
            metriques = calculer_metriques_avancees(rendements_action, rendements_indice_filtres, prix_action)

            if metriques:
                # Ajouter les résultats
                nouvelle_ligne = pd.DataFrame([{
                    'Secteur': secteur,
                    'Entreprise': entreprise,
                    'Symbole': symbole,
                    'Alpha': metriques['Alpha'],
                    'Beta': metriques['Beta'],
                    'R-squared': metriques['R-squared'],
                    'Correlation': metriques['Correlation'],
                    'Rendement_Geo_Annualisé': metriques['Rendement_Geo_Annualisé'],
                    'Volatilité_Totale': metriques['Volatilité_Totale'],
                    'Volatilité_Systématique': metriques['Volatilité_Systématique'],
                    'Volatilité_Résiduelle': metriques['Volatilité_Résiduelle']
                }])

                resultats = pd.concat([resultats, nouvelle_ligne], ignore_index=True)
            else:
                print(f"  Impossible de calculer les métriques pour {entreprise}")

        # Arrondir les résultats numériques à 4 décimales
        colonnes_numeriques = ['Alpha', 'Beta', 'R-squared', 'Correlation', 'Rendement_Geo_Annualisé',
                               'Volatilité_Totale', 'Volatilité_Systématique', 'Volatilité_Résiduelle']
        resultats[colonnes_numeriques] = resultats[colonnes_numeriques].round(4)

        # Calculer également les moyennes par secteur
        print("\nCalcul des moyennes par secteur...")
        moyennes_secteur = resultats[resultats['Secteur'] != "Indice"].groupby('Secteur')[
            colonnes_numeriques].mean().round(4)
        moyennes_secteur['Entreprise'] = 'MOYENNE'
        moyennes_secteur['Symbole'] = '-'
        moyennes_secteur = moyennes_secteur.reset_index()

        # Réorganiser les colonnes des moyennes pour qu'elles correspondent aux résultats
        moyennes_secteur = moyennes_secteur[colonnes_resultats]

        # Ajouter les moyennes aux résultats
        resultats_complets = pd.concat([resultats, moyennes_secteur], ignore_index=True)

        # Réordonner pour que l'indice soit en premier
        resultats_complets = pd.concat([
            resultats_complets[resultats_complets['Secteur'] == "Indice"],
            resultats_complets[resultats_complets['Secteur'] != "Indice"]
        ]).reset_index(drop=True)

        # Créer également un tableau de métriques basiques (pour le code 2)
        metriques_basiques = resultats_complets[['Symbole', 'Alpha', 'Beta', 'R-squared',
                                                 'Rendement_Geo_Annualisé', 'Volatilité_Totale']]
        metriques_basiques = metriques_basiques.set_index('Symbole')

        # Exporter les résultats
        try:
            # Exporter l'analyse sectorielle complète
            with pd.ExcelWriter('analyse_sectorielle_5ans.xlsx') as writer:
                # Exporter tous les résultats sur une feuille
                resultats_complets.to_excel(writer, sheet_name='Analyse Complète', index=False)

                # Exporter les détails par secteur sur des feuilles séparées
                secteurs_uniques = [s for s in resultats_complets['Secteur'].unique() if s != "Indice"]
                for secteur in secteurs_uniques:
                    secteur_df = resultats_complets[resultats_complets['Secteur'] == secteur]
                    if not secteur_df.empty:
                        secteur_df.to_excel(writer, sheet_name=secteur[:31],
                                            index=False)  # Excel limite les noms de feuilles à 31 caractères

                # Exporter les moyennes par secteur
                moyennes_secteur.to_excel(writer, sheet_name='Moyennes par Secteur', index=False)

                # Exporter l'indice séparément
                resultats_complets[resultats_complets['Secteur'] == "Indice"].to_excel(writer, sheet_name='Indice',
                                                                                       index=False)

            # Exporter également les métriques basiques (résultat du code 2)
            with pd.ExcelWriter('metriques_performance_5ans.xlsx') as writer:
                metriques_basiques.to_excel(writer, sheet_name='Métriques Performance')

            print("\nAnalyse sectorielle exportée avec succès dans 'analyse_sectorielle_5ans.xlsx'!")
            print("Métriques basiques exportées dans 'metriques_performance_5ans.xlsx'!")

        except ImportError:
            resultats_complets.to_csv('analyse_sectorielle_5ans.csv', index=False)
            metriques_basiques.to_csv('metriques_performance_5ans.csv')
            print("\nRésultats exportés avec succès au format CSV!")
            print("Pour sauvegarder en Excel, installez openpyxl: pip install openpyxl")

        print("\nAnalyse complète sur 5 ans terminée avec succès!")
        print(f"Nombre d'actions analysées: {len(resultats) - 1}")  # -1 pour l'indice

    except Exception as e:
        print(f"Erreur lors de l'exécution de l'analyse: {e}")


if __name__ == "__main__":
    main()