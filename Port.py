import pandas as pd
import yfinance as yf
import datetime
import time

# Définir la période d'analyse (5 ans)
date_fin = datetime.datetime.now().strftime('%Y-%m-%d')
date_debut = (datetime.datetime.now() - datetime.timedelta(days=5*365)).strftime('%Y-%m-%d')

print(f"Période d'analyse: du {date_debut} au {date_fin}")

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

# Créer une liste de symboles à partir du mapping
symboles = list(symbole_mapping.values())

# Dictionnaire de taux de change fixes
taux_fixes = {
    'GBP': 1.15,  # 1 GBP = 1.15 EUR
    'CHF': 0.95,  # 1 CHF = 0.95 EUR
    'SEK': 0.086, # 1 SEK = 0.086 EUR
    'DKK': 0.134, # 1 DKK = 0.134 EUR
    'NOK': 0.086, # 1 NOK = 0.086 EUR
    'USD': 0.85,  # 1 USD = 0.85 EUR
    'EUR': 1.0,   # déjà en euros
}

# Dictionnaire de correspondance pour les devises selon la bourse
devises_par_bourse = {
    '.L': 'GBP',    # Londres - Livre sterling
    '.SW': 'CHF',   # Suisse - Franc suisse
    '.BR': 'EUR',   # Bruxelles - Euro
    '.PA': 'EUR',   # Paris - Euro
    '.AS': 'EUR',   # Amsterdam - Euro
    '.DE': 'EUR',   # Allemagne - Euro
    '.MC': 'EUR',   # Madrid - Euro
    '.ST': 'SEK',   # Stockholm - Couronne suédoise
    '.MI': 'EUR',   # Milan - Euro
    '.CO': 'DKK',   # Copenhague - Couronne danoise
    '.OL': 'NOK',   # Oslo - Couronne norvégienne
    '.VI': 'EUR',   # Vienne - Euro
    '.LS': 'EUR',   # Lisbonne - Euro
    '.HE': 'EUR',   # Helsinki - Euro
    '.I': 'EUR',    # Irlande - Euro
}

# Fonction pour déterminer la devise d'un symbole
def obtenir_devise(symbole):
    for extension, devise in devises_par_bourse.items():
        if symbole.endswith(extension):
            return devise
    return 'EUR'  # Par défaut, utiliser l'Euro

# Fonction pour récupérer les données historiques avec gestion des erreurs
def obtenir_donnees_historiques(symbole, date_debut, date_fin, tentatives_max=3, delai=1):
    for tentative in range(tentatives_max):
        try:
            stock = yf.Ticker(symbole)
            historique = stock.history(start=date_debut, end=date_fin)
            if not historique.empty:
                # Supprimer les informations de fuseau horaire
                historique.index = historique.index.tz_localize(None)
                return historique
            time.sleep(delai)
        except Exception as e:
            print(f"Erreur pour {symbole} (tentative {tentative+1}): {e}")
            time.sleep(delai)

    # Si toutes les tentatives échouent, retourner un DataFrame vide
    print(f"Impossible de récupérer les données pour {symbole} après {tentatives_max} tentatives")
    return pd.DataFrame()

# Définir le symbole pour l'indice STOXX 600
indice_eurostoxx = "^STOXX"  # Symbole correct pour Yahoo Finance

# Récupérer d'abord les données de l'indice STOXX 600
print(f"Récupération des données pour l'indice {indice_eurostoxx}...")
historique_indice = obtenir_donnees_historiques(indice_eurostoxx, date_debut, date_fin)

if historique_indice.empty:
    print(f"Impossible de récupérer les données pour {indice_eurostoxx}. Utilisation d'un indice alternatif.")
    # Essayer avec un autre symbole
    indice_eurostoxx = "SX5E.PA"  # Autre tentative avec le STOXX50E
    historique_indice = obtenir_donnees_historiques(indice_eurostoxx, date_debut, date_fin)

    if historique_indice.empty:
        print(f"Impossible de récupérer les données pour {indice_eurostoxx}. Création d'un indice synthétique.")
        # Créer un indice synthétique à partir des actions
        prix_actions = pd.DataFrame()

        for i, symbole in enumerate(symboles[:10]):  # On prend les 10 premières actions pour l'indice synthétique
            print(f"Récupération des données pour l'indice synthétique - {symbole}... ({i+1}/10)")
            historique = obtenir_donnees_historiques(symbole, date_debut, date_fin)

            if not historique.empty:
                if prix_actions.empty:
                    prix_actions = pd.DataFrame(index=historique.index)
                prix_actions[symbole] = historique['Close']

        if not prix_actions.empty:
            # Normaliser chaque série à 100 à la première date
            prix_actions = prix_actions.dropna(axis=1, thresh=len(prix_actions)*0.7)
            normalise = prix_actions.div(prix_actions.iloc[0]).mul(100)

            # Créer l'indice synthétique
            historique_indice = pd.DataFrame(index=normalise.index)
            historique_indice['Close'] = normalise.mean(axis=1)
            indice_eurostoxx = "INDICE_SYNTHETIQUE"
            print("Indice synthétique créé avec succès")
        else:
            print("Impossible de créer un indice synthétique. Arrêt du programme.")
            exit()
else:
    print(f"Données récupérées avec succès pour {indice_eurostoxx}")

# Créer un dictionnaire pour stocker toutes les données de prix
donnees_prix = {}

# Ajouter l'indice
donnees_prix[indice_eurostoxx] = historique_indice['Close']

# Récupérer les données pour les actions sélectionnées et les convertir en euros
for i, symbole in enumerate(symboles):
    print(f"Récupération des données pour {symbole}... ({i+1}/{len(symboles)})")

    # Récupération des données historiques
    historique = obtenir_donnees_historiques(symbole, date_debut, date_fin)

    if not historique.empty:
        # Déterminer la devise
        devise = obtenir_devise(symbole)

        # Utiliser un taux de change fixe pour la conversion
        if devise != 'EUR':
            print(f"  Conversion de {devise} vers EUR pour {symbole} (taux fixe)")
            taux = taux_fixes.get(devise, 1.0)
            donnees_prix[symbole] = historique['Close'] * taux
        else:
            # Si déjà en euros, ajouter directement
            donnees_prix[symbole] = historique['Close']

    # Pause pour éviter les limitations d'API
    time.sleep(0.25)

# Créer le DataFrame des prix de clôture à partir du dictionnaire (évite la fragmentation)
print("\nCréation du DataFrame des prix de clôture...")
prix_cloture = pd.DataFrame(donnees_prix)

# Remplir les valeurs manquantes avec la valeur précédente
prix_cloture.fillna(method='ffill', inplace=True)

# Calculer les rendements journaliers en une seule opération
print("Calcul des rendements journaliers...")
rendements_journaliers = prix_cloture.pct_change(fill_method=None)  # Évite le FutureWarning

# Nettoyage et tri des données
prix_cloture = prix_cloture.sort_index()
rendements_journaliers = rendements_journaliers.sort_index()

# Suppression des lignes sans données
prix_cloture = prix_cloture.dropna(how='all')
rendements_journaliers = rendements_journaliers.dropna(how='all')

# Création d'une copie avec les dates au format jj/mm/aaaa
prix_cloture_date_format = prix_cloture.copy()
prix_cloture_date_format.index = prix_cloture_date_format.index.strftime('%d/%m/%Y')
rendements_journaliers_date_format = rendements_journaliers.copy()
rendements_journaliers_date_format.index = rendements_journaliers_date_format.index.strftime('%d/%m/%Y')

# Exporter les résultats
try:
    # Exporter en Excel
    with pd.ExcelWriter('resultats_actions_5ans.xlsx') as writer:
        prix_cloture_date_format.to_excel(writer, sheet_name='Prix Quotidiens')
        rendements_journaliers_date_format.to_excel(writer, sheet_name='Rendements Journaliers')
    print("\nDonnées exportées avec succès au format Excel!")
except ImportError:
    # Exporter en CSV si openpyxl n'est pas disponible
    prix_cloture_date_format.to_csv('prix_quotidiens_5ans.csv')
    rendements_journaliers_date_format.to_csv('rendements_journaliers_5ans.csv')
    print("\nDonnées exportées avec succès au format CSV!")
    print("Pour sauvegarder en Excel, installez openpyxl: pip install openpyxl")

print(f"Données extraites sur 5 ans pour {len(symboles)} actions et 1 indice.")
print("Traitement des données terminé!")
