#pip install -r requirements.txt
#python -m pip install --upgrade pip
#pip install --upgrade pip
#pip install --upgrade streamlit
import streamlit as st 
import pandas as pd
from datetime import datetime, timedelta
import pytz
import re

#import os
#os.chdir(r"D:\Dashboard_B2B\form_streamlit\dashbord_ajuste")
##os.chdir(r"C:\Users\HZXH9786\Downloads\old_pilotage_b2b")


url = "C:/Users/DTDS3300/Desktop/Orange/Fichier_Formulaire/base_dashboard.xlsx"
commerciaux_cpt = pd.read_excel(url, sheet_name="commerciaux_cpt")
commerciaux_sgt = pd.read_excel(url, sheet_name="commerciaux_sgt")
Support = pd.read_excel(url, sheet_name="Support")
type_vente = pd.read_excel(url, sheet_name="type_vente")
produit_item = pd.read_excel(url, sheet_name="produit_item")




    
def main():
    st.header("REPORTING DE SUIVI DES VENTES VERTICAUX")
    #st.write("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    st.title("Infos:")
    
    st.markdown(
        ("\nPour être performant et réussir dans ses missions,\
         Il est nécessaire de consolider les objectifs. Pour cela il faut :\
             \n**Compréhension de la portée du travail** \
                 \n**Compréhension du business et de la vision de l'entreprise** \
                 \n**Compréhension du système d’évaluation** \
                     \n**Reporter fidèlement ses résultats**")
        )
    st.write("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
    #st.subheader('Auteur: achillefomekong97@gmail.com')
    with st.form(key="pilotage_form"):
        def nettoyer_chaine(chaine):
            if chaine is not None:
                chaine1=str(chaine)
                chaine_nettoyee = chaine1.strip()
                return chaine_nettoyee if chaine_nettoyee else None
            else: return None
            
        def val_filtre(df,val, col_filtre, col):
            l=pd.Series(df[col].apply(lambda x:str(x)).unique()).sort_values().tolist()
            if val is not None:       
                b=df[col_filtre].apply(lambda x:str(x))==str(val)
                l=pd.Series(df[b][col].apply(lambda x:str(x)).unique()).sort_values().tolist()
                return l
            
        signal=1
        st.sidebar.title("Infos du commercial et produit:")
        commercial=st.sidebar.selectbox('Nom et prénom du commercial**',pd.Series(commerciaux_cpt['Commerciaux'].unique()).sort_values().tolist())
        
        password = nettoyer_chaine(st.sidebar.text_input("Mot de passe", type="password", max_chars=4))
        signal22=0
        signal23=1
        if password is not None and len(password) == 4:
            signal22=1
        else:
            st.sidebar.warning("Le mot de passe doit contenir  4 caractères non tous vides.")
            
        if signal22==1:
            verif=val_filtre(commerciaux_sgt, commercial,'Commerciaux', 'Password')
            if password != verif[0]:
                st.sidebar.warning("Mot de passe incorrect")
            else: signal23=2
            
            
        if signal23==2:    
            
            
        
              l1=val_filtre(commerciaux_sgt, commercial,'Commerciaux', 'Segment')
              segment=l1[0]
              #segment=st.sidebar.selectbox('Segment**',l1)  
        
              l11=val_filtre(Support,commercial,"Commerciaux", 'SUPPORTS')
              support=l11[0]
              #support=st.sidebar.selectbox('Support**',l11)
        
        
              #st.sidebar.write('<span style="color: orange;">Segment:'+segment'</span>', unsafe_allow_html=True)
              st.sidebar.write('<span style="color: orange;">Segment: ' + str(segment) + '</span>', unsafe_allow_html=True)
              st.sidebar.write('<span style="color: orange;">Support:: ' + str(support) + '</span>', unsafe_allow_html=True)

              #st.sidebar.write(" Support:"+support)
            

              produit=st.sidebar.selectbox('Produit**',pd.Series(produit_item['PRODUITS'].unique()).sort_values().tolist())
            
              l2=val_filtre(produit_item, produit,'PRODUITS', 'ITEMS')
              item=st.sidebar.selectbox('Item**',l2)
            
              l23=["JPO ONE DAY AT CUSTOMER","MUTUELLES/CORPORATIONS PROSPECTION",\
                  "CODIR 01 MOIS – 01 SECTEUR D’ACTIVITÉ","CAMPAGNE PHONING",\
                      "TIRS GROUPÉS", "DÉPÔT COURRIERS", "AUTRES ACTIONS IMPACTANTES",\
                          "BUSINESS REVIEW"]
              l23=sorted(l23)
              action=st.sidebar.selectbox('Action ayant conduit à la vente**',l23)
        
              if produit=="LL" or produit=="ICT/MESSAGING PRO":
                  date_actuelle = datetime.now().date()
                  date_livraison = st.sidebar.date_input("Date de livraison: \(format=année/mois/jour/)", value=date_actuelle, max_value=date_actuelle)
                  date_livraison = date_livraison.strftime("%d-%m-%Y")
              else: date_livraison="______"
        
        
       
  
              st.sidebar.title("Informations client:")
              l3=val_filtre(commerciaux_cpt, commercial,'Commerciaux', 'Client principal')
              client=st.sidebar.selectbox("Client correspondant** \n \
                                           (selectionner * **Autre** * au cas ou ce client n\'est pas precise\)",l3+['_Autre'])
              def nettoyer_chaine(chaine):
                  if chaine is not None:  
                      chaine1=str(chaine)
                      chaine_nettoyee = chaine1.strip()
                      return chaine_nettoyee if chaine_nettoyee else None
                  else: return None
              #commerciaux_cpt['Compte Principal']=commerciaux_cpt['Compte Principal'].apply(nettoyer_chaine)
              if client=='_Autre':
                  st.sidebar.write("++++ **Section reservée à cet autre compte** ++++")
                  autre_client=st.sidebar.text_input("Préciser ce nom du client correspondant**", value=None)
                  acct_nbr=st.sidebar.text_input("Préciser son numero de compte**", value=None)
                  def est_format_valide(ch):
                      pattern =  r'^[1-4]\.\d+(\.\d+)?$'
                      return bool(re.match(pattern, ch))
                  if nettoyer_chaine(acct_nbr) is not None:                
                      if est_format_valide(nettoyer_chaine(acct_nbr))==False:
                          st.sidebar.write('<span style="color: red;">Format invalide \(4.nnnnnn...\)</span>', unsafe_allow_html=True)
                          signal=-1
                          st.sidebar.write("++++++++++++++++++++++++++++++++++++++++++++")
              else: 
                  autre_client='______'
                  l4=val_filtre(commerciaux_cpt, client,'Client principal', 'Compte Principal')
                  #acct_nbr=st.sidebar.selectbox('Compte correspondant**',l4)
                  acct_nbr=l4[0]
                  st.sidebar.write('<span style="color: orange;">Numero de compte: ' + str(acct_nbr) + '</span>', unsafe_allow_html=True)

     
              statut=st.sidebar.selectbox('Statut du compte**',['Ancien', 'Nouveau'])
    
              st.sidebar.title("Facturation:")
              qte=st.sidebar.number_input('Quantité**', min_value=1, value=1)
        
              ttc_ht=st.sidebar.selectbox('TTC ou Hors taxe?**',['TTC', 'HT'])
        
              ###pu = st.sidebar.number_input("Prix unitaire**", step=0.01, format="%.2f")
    
              ###prix_total=qte*pu
              #st.sidebar.write("Prix Total= "+str(prix_total))
              ###st.sidebar.write('<span style="color: orange;">Prix Total:: ' + str(prix_total) + '</span>', unsafe_allow_html=True)
            
              if ttc_ht=="TTC":
                  mt_ttc = st.sidebar.number_input("Prix Total TTC**", step=0.01, format="%.2f")
                  mt_ht = None
              elif ttc_ht=="HT":
                  mt_ht= st.sidebar.number_input("Prix Total HT**", step=0.01, format="%.2f")
                  mt_ttc=None

            
              typ_vente=st.sidebar.selectbox('Type de vente**',pd.Series(type_vente['TYPE DE VENTE'].unique()).sort_values().tolist())
              num_ticket=st.sidebar.text_input("Numéro de ticket**", value=None)

              def est_format_valide_2(chaine):
                  motif = r'^CAS-\d{8}-[A-Z0-9]{6}$'
                  if re.match(motif, chaine):
                      return True
                  else:
                      return False
            


              #time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            
              ##gestion du fuseau horaire.
              # Obtenir le fuseau horaire du Cameroun
              cameroon_tz = pytz.timezone('Africa/Douala')
              # Obtenir le temps actuel dans le fuseau horaire du Cameroun
              time = datetime.now(cameroon_tz).strftime("%d-%m-%Y %H:%M:%S")
              val = pd.to_datetime(time, format="%d-%m-%Y %H:%M:%S")
              week_number = val.strftime('%V')
              week_number="S"+week_number
                   
            
            
              commercial=nettoyer_chaine(commercial)
              produit=nettoyer_chaine(produit)
              item=nettoyer_chaine(item)
              statut=nettoyer_chaine(statut)
              acct_nbr=nettoyer_chaine(acct_nbr)
              autre_client=nettoyer_chaine(autre_client)
              client=nettoyer_chaine(client)
              num_ticket=nettoyer_chaine(num_ticket)
              ttc_ht=nettoyer_chaine(ttc_ht)
              typ_vente=nettoyer_chaine(typ_vente)
              data={
                  'Honorateur':time,
                  'Nom commercial':commercial,
                  'Segment': segment,
                  'Produit': produit,
                  'Item produit':item,
                  'Date Livraison':date_livraison,
                  'Support':support,
                  'Numero compte': acct_nbr,
                  'Autre Client': autre_client,
                  'Client': client,
                  'Statut du compte': statut,
                  'Quantite': qte, 
                  'TTC/HT':ttc_ht,
                  'Prix Total TTC': mt_ttc,
                  'Prix Total HT': mt_ht,
                
                  'Prix unitaire': "End",
                  'Prix Total': "End",
                  'Numero Ticket': num_ticket,
                  'Type Vente': typ_vente,
                  'Semaine': week_number,
                  'Action conduisant à la vente':action
                  }

            
              #c1=((produit=="EQUIPEMENT/DE/FMS") and (typ_vente!="LEASING")) or (produit !="EQUIPEMENT/DE/FMS")
              c2=produit in {"EQUIPEMENT/DE/FMS","ICT/MESSAGING PRO","OM", "LL"}
              #c3=produit != "CAS VOIX\GSM\ACTIVATION"
              #c=c1 or c2 or c3
              if (nettoyer_chaine(num_ticket) is not None):
                  #if c1:
                  if not c2:
                      if est_format_valide_2(nettoyer_chaine(num_ticket))==False:
                          st.sidebar.write('<span style="color: red;"> Format invalide pour ce produit ou ce type de vente. Le format attendu est: CAS-XXXXXXXX-ZZZZZZ, où X est un chiffre et Z est une lettre majuscule ou un chiffre</span>', unsafe_allow_html=True)
                          signal=-2


            
            
              df=pd.DataFrame(data, index=[0])
              df["Numero compte"] = df["Numero compte"].astype(str)
              st.dataframe(df)
              #st.write("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
              st.write('<span style="color: orange;">Veuillez vérifier le tableau avant de cliquer sur \"Envoyer\". Après la validation, un message de succès apparaitra.</span>', unsafe_allow_html=True)
                       #st.write("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                       #sign=1
              submit_button=st.form_submit_button(label="Envoyer le formulaire")
              #signal=1
              if client=='_Autre':
                  valeur=[commercial, segment, produit, item, support, acct_nbr, autre_client, client, statut, qte,ttc_ht, \
                          num_ticket, typ_vente]
              else: valeur=[commercial, segment, produit, item, support, acct_nbr, client, statut, qte,ttc_ht, \
                        num_ticket, typ_vente]
        
              if submit_button:
                  for val in valeur:
                      if val is None:
                          signal=0
                          break
                    
             # def est_format_valide(ch):
              #    pattern = r'^[1-9]\.\d+$'
              #    return bool(re.match(pattern, ch))
             # if est_format_valide(acct_nbr)==False:
               #   signal=-1

                  if signal==0:
                      st.write('<span style="color: red;">Error: SVP, veuillez remplir d\'abord tous les champs réquis (suivis de **) </span>', unsafe_allow_html=True)
                  elif signal==-1:  
                      st.write('<span style="color: red;">Error: SVP, veuillez respecter le format des numeros des comptes principaux</span>', unsafe_allow_html=True)
                  elif signal==-2:  
                      st.write('<span style="color: red;">Error: SVP, veuillez respecter le format des numeros de tickets</span>', unsafe_allow_html=True)
                
                  else: 
                      import gspread
                      from gspread_dataframe import get_as_dataframe, set_with_dataframe
                      sa = gspread.service_account(filename="projet-pilotage-812e389beb7c.json")
                      sh = sa.open("pilotage_business")
                      worksheet = sh.worksheet("base")
                      # Convertir le DataFrame en une structure de données que gspread comprend
                      existing_df = get_as_dataframe(worksheet, evaluate_formulas=True, dtype=str)
                      #st.dataframe(existing_df)
                    
                      #existing_df["Numero compte"] = existing_df["Numero compte"].astype(str)
                    
                      existing_df = existing_df.iloc[:, 0:21] ## les 21 premieres colonnes  
                      #st.dataframe(existing_df)
                      # Supprimer les lignes où la colonne A est vide
                      existing_df = existing_df[existing_df.iloc[:, 1].notna()]
                      #st.dataframe(existing_df)
                      ####existing_tickets = ['__']+existing_df['Numero Ticket'].astype(str).tolist()
                    
                      existing_acct = commerciaux_cpt['Compte Principal'].astype(str).tolist()
                    
                      #if existing_df['Numero Ticket'].astype(str).eq(str(num_ticket)).any():
                      ###if str(num_ticket) in existing_tickets:
                          ###date_existence = existing_df.loc[existing_df['Numero Ticket'].astype(str) == str(num_ticket), 'Honorateur'].values[0]
                          ###date_existence = datetime.strptime(date_existence, "%d-%m-%Y %H:%M:%S")
                          ###formatted_date = date_existence.strftime('%d/%m/%Y %H:%M:%S')
                          ###st.write(f'<span style="color: red;">Error: ce numéro de ticket existe déjà depuis le {formatted_date}</span>', unsafe_allow_html=True)
                          ###st.stop()
                      if client=='_Autre annulé' and str(acct_nbr) in existing_acct:
                          client_exist = commerciaux_cpt.loc[commerciaux_cpt['Compte Principal'].astype(str) == str(acct_nbr), 'Commerciaux'].values[0]
                          nom = commerciaux_cpt.loc[commerciaux_cpt['Compte Principal'].astype(str) == str(acct_nbr), 'Client principal'].values[0]
                          st.write(f'<span style="color: red;">Error: ce numero de compte (dont le client est {nom}) appartient au commercial {client_exist}</span>', unsafe_allow_html=True)
                          st.write(f'<span style="color: red;"> Veuillez consulter votre portefeuille client dans la liste deroulante pour selectionner le client ou, assurez-vous que ce client vous appartient ou\
                                        que le numero de compte soit exact. Une possibilité aussi est qu\'il ne soit pas un autre client comme selectionné. </span>', unsafe_allow_html=True)
                          st.stop()
                    
                      else:
                          # Ajout de nouvelles lignes au DataFrame existant
                        
                          updated_df = pd.concat([existing_df, df], ignore_index=True)
                          #updated_df["Numero compte"] = updated_df["Numero compte"].astype(str)
                          #st.dataframe(updated_df.dtypes)
                          #st.dataframe(updated_df)
                        
                          # Mettre à jour la feuille Google Sheets avec le DataFrame mis à jour
                          set_with_dataframe(worksheet, updated_df, resize=True)
                        
                          st.success("Formulaire envoyé avec succes")
                          #st.stop()
              st.write("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
              affiche_button=st.form_submit_button(label="Afficher vos données")

              default_end_date = datetime.now()
              default_start_date = default_end_date - timedelta(days=6)

              start_date = st.date_input("Début de la période", value=default_start_date, max_value=default_end_date)
              end_date = st.date_input("Fin de la période", value=default_end_date, min_value=start_date, max_value=default_end_date)+ timedelta(days=1)

              start_date_str = start_date.strftime("%d-%m-%Y %H:%M:%S")
              end_date_str = end_date.strftime("%d-%m-%Y %H:%M:%S")
            
              start_date_str=pd.to_datetime(start_date_str, format="%d-%m-%Y %H:%M:%S")
              end_date_str=pd.to_datetime(end_date_str, format="%d-%m-%Y %H:%M:%S")
            
              #st.write(start_date_str)
              #st.write(end_date_str)


              if affiche_button:
                  import gspread
                  from gspread_dataframe import get_as_dataframe, set_with_dataframe
            
                  sa1 = gspread.service_account(filename="projet-pilotage-812e389beb7c.json")
                  sh1 = sa1.open("pilotage_business")
                  worksheet1 = sh1.worksheet("base")
                  # Convertir le DataFrame en une structure de données que gspread comprend
                  existing_df1 = get_as_dataframe(worksheet1, evaluate_formulas=True, dtype=str)
                  existing_df1 = existing_df1.iloc[:, 0:19] #
                  boo=existing_df1['Nom commercial']==commercial
                
                  existing_df1['Honorateur'] = pd.to_datetime(existing_df1['Honorateur'], format="%d-%m-%Y %H:%M:%S")
                
                  filtered_df1 = existing_df1[(existing_df1['Honorateur'] >= start_date_str) & (existing_df1['Honorateur'] <= end_date_str) & boo]

                  #st.write("Résultats filtrés :")
                  st.dataframe(filtered_df1.sort_values(by='Honorateur', ascending=False))
                  #st.write(filtered_df)
            
                            
              
    
if __name__=='__main__':
    main()
    
