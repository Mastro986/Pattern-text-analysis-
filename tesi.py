import numpy as np 
import docx
import pandas as pd
#libreria che mi servirà per le tabelle
from tabulate import tabulate
#creo una funzione che dato un file mi restituisce il contenuto del testo
def getText(filename): 
    doc = docx.Document(filename) 
    fullText = [] 
    for para in doc.paragraphs: 
        fullText.append(para.text) 
    return '\n'.join(fullText)
#funzione che controlla la punteggiatura e tronca la parola dove c'è la punteggiatura--> restituisce la parola troncata
def check_punctuation(frase):
    parola_troncata=""
    parola=frase.split()
    for p in parola:
        if p.__contains__("(") or p.__contains__(")") or p.__contains__("{") or p.__contains__("}") or p.__contains__("[") or p.__contains__("]") or p.__contains__("/") or p.__contains__("%") or p.__contains__("&") or p.__contains__("°") or p.__contains__("^") or p.__contains__("|") or p.__contains__("!") or p.__contains__("?") or p.__contains__("<") or p.__contains__(">") or p.__contains__("."):
            return parola_troncata
        #controlla se la parola abbia un numero al suo interno
        if p.__contains__("0") or p.__contains__("1") or p.__contains__("2") or p.__contains__("3") or p.__contains__("4") or p.__contains__("5") or p.__contains__("6") or p.__contains__("7") or p.__contains__("8") or p.__contains__("9"):
            return parola_troncata
        if p.__contains__(";") or p.__contains__(",") :
            parola_troncata=parola_troncata+" "+p.replace(',', '')
            return parola_troncata
        else:
            parola_troncata=parola_troncata+" "+p
    return parola_troncata
#funzione che data una frase ritorna il numero di parole che la contengono
def tell_word(frase):
    i=0
    words=frase.split()
    for word in words:
        i=i+1
    return i
#funzione che controlla gli articoli e i complementi appartenenti ad una parola 
def check_word(frase):
    parola=""
    words=frase.split()
    for word in words:
        if(len(word)==1):
            parola=""
            return parola
        if(word=="Le" or word=="La" or word=="L'" or word=="Ii" or word=="Ie" or word=="Li" or word=="Ja" or word=="Les" or word=="Un" or word=="Une" or word=="Des" or word=="A" or word=="De" or word=="Avec" or word=="En" or word=="U" or word=="Notre" or word=="Votre" or word=="Leur" or word=="Nostres" or word=="Nostre" or word=="Mon" or word=="Mes" or word=="Ton" or word=="Son" or word=="Jj" or word=="Par" or word=="Por" or word=="Se" or word=="Ne" or word=="El" or word=="Qu’a" or word=="Mais" or word=="C’est" or word=="Pour" or word=="Per" or word=="An" or word=="Ou" or word=="Noustre"):
            parola=""
            return parola
        else: 
            parola=parola+" "+word 
    return parola 
#funzione che ricava nome ,cognome e contesto 
def getScore(filename,name_text):
    lista=[]
    lista.extend([[name_text," "," "]])
    #variabile che mi servirà per i contesti
    contesto=""
    #variabile min
    min=0
    #contatore indice di ogni parola
    i=0
    #variabile che andrà a contenere nome e cognome temporanei
    Score=""
    #variabile che andrà a contenere tutti i nomi e i cognomi
    ScoreTot=""
    #contatore 
    contatore=0
    #stringa per controllo virgole
    obbiettivo=""
    #mi ricavo il testo e lo metto dentro a questa variabile stringa con la funzione getText()
    testo_completo=getText(filename)
    #divide la stringa(testo_completo) in sottostringhe
    words = testo_completo.split()
    #numero parole che contiene la stringa-->mi servirà successivamente per non andare fuori range nei controlli
    numero_parole=tell_word(testo_completo)
    for word in words: 
        
        #verifica se la prima lettera della parola estratta è maiuscola 
        if(word[0].isupper() and word.isupper()==False and word!="Et" and word!="Que"):
            Score=Score+" "+word
            contatore=contatore+1
            #controllo di stare nei pressi del range
            if((i+1)<numero_parole):
                #se ci sono 2 o più parole consecutive con la lettera maiuscola e se la parola successiva è scritta in minuscolo allora aggiungo a scoreTot la parola 
                if(contatore>1 and words[i+1].islower()==True):
                    #controllo punteggiatura e numeri all'interno delle parole e tronca le parole che contengono questi elementi
                    Score=str(check_punctuation(Score))
                    #controllo articoli e complementi
                    Score=check_word(Score) 
                    #funzione che controlla la lunghezza di Score, ovvero della parola-->se la parola è una non la considero perchè io voglio una struttura del tipo[nome,cognome]
                    n_word=tell_word(Score)
                    if(n_word<=1):
                        Score=""
                    else:
                        #ScoreTot=ScoreTot+" "+Score+"      "
                        ScoreTot=Score
                        #prendo il contesto, considero 2 parole prima e 2 parole dopo la frase
                        min=i-n_word-19
                        while min<=i+20:
                            if(min<numero_parole):
                                contesto=contesto+" "+words[min]
                                min=min+1
                            else:
                                break
                        #aggiungo alla lista
                        lista.extend([[" ",ScoreTot,contesto]])
                      
            #se invece la parola successiva alle precedenti è maiuscola non faccio nulla e aggiungo a scoretot la parola completa quando troverò una successiva tutta minuscola
        else:
            contatore=0
            Score=" "
            contesto=""
            min=0
        
        i=i+1
           
    return lista      
  
#...................................
#...................................

#variabili che contengono il nome del testo
nome_f1="Balain, roman de, Dominica Legge.docx"
nome_f2="Barlaam et Josaphat, Chardry.docx"
nome_f3="Barlaam et Josaphat, Gui de Cambrai, Appel.docx"
nome_f4="Barlaam et Josaphat, Mills.docx"
nome_f5="Barlaam et Josaphat, monte Athos, Agrigoraei.docx"
nome_f6="Barlaam et Josaphat, Sonet, t. I.docx"
nome_f7="Barlaam et Josaphat, Sonet.docx"
nome_f8="Baudouin de Flandre, Pinto-Mathieu.docx"
nome_f9="Bel Inconnu, Perret Weil.docx"
nome_f10="Belle Helene de Costantinople, Jean Wauquelin, Crecy 2022.docx"
nome_f11="Belle Helene de Costantinople, Roussel.docx"
nome_f12="Bérinus, prosa, Boussat.docx"
nome_f13="Bérinus, versi, Bossaut.docx"
nome_f14="Berta de li gran pié, Morgan.docx"
nome_f15="Berte, histoire de la reine, Tylus.docx"
nome_f16="Bertes aus grans piés, Adenet le Roy, Henry.docx"
nome_f17="Biaodouz, Lemaire.docx"
nome_f18="Blancandin, en prose, Greco.docx"
nome_f19="Brun de la Montagne.docx"
nome_f20="Brut, en prose.docx"
nome_f21="Brut, roman de, Wace, Arnold.docx"

nome_f1c="Cardenois, roman de, Cocco.docx"
nome_f2c="Cassidorus, roman de, t. I, Palermo.docx"
nome_f3c="Cassidorus, roman de, t. II, Palermo.docx"
nome_f4c="Charles de Hongrie, roman de, Chenerie.docx"
nome_f5c="Charrette, conte de, Combes.docx"
nome_f6c="chastelain de Vergi, Stuip.docx"
nome_f7c="Chastelaine de Coucy, Babbi.docx"
nome_f8c="Chatelain de Coucy et dame de Fayel, roman de, Gaullier-Bougassas.docx"
nome_f9c="Cheval de Fust, Meliacin, ed. Saly.docx"
nome_f10c="Chevalerie de Judas Macabée, Smeets.docx"
nome_f11c="Chevalier à l_épée, Johnston-Owen.docx"
nome_f12c="Chevalier as deus espees, le, Rockwell.docx"
nome_f13c="Chevalier au papegau, conte de, Charpentier-Victorin.docx"
nome_f14c="Ciperis de Vignevaux, Ramello.docx"
nome_f15c="Claris et Laris, Pierreville.docx"
nome_f16c="Clarisse et Florent.docx"
nome_f17c="Cleomades, en prose, Trachsler Mailliet.docx"
nome_f18c="Cleomadés, Henry.docx"
nome_f19c="Cleriadus et Meliadice, Zink.docx"
nome_f20c="Cligés, Chretien de Troyes, Foerster.docx"
nome_f21c="Cligés, en prose, Colombo-Timelli.docx"
nome_f22c="Comte d_Artois, roman de, Seignereut.docx"
nome_f23c="comte de Poitiers, roman de, Malmerg.docx"
nome_f24c="Cristal et Clarie, Breuer.docx"
nome_f25c="Croissant, roman de, Schafer.docx"

nome_f1d="063 Dame au liocorne, roman de.docx"
nome_f2d="064 Didot-Perceval, Roach.docx"
nome_f3d="065 Durmart le Galois, Gildea INTERNET.docx"

nome_f1e="066 Eledus et Serene, roman de, Reinall.docx"
nome_f2e="067 Eneas, roman de, Salverda de Grave.docx"
nome_f3e="068 Eneas, roman de, t. II, Salverda de Grave.docx"
nome_f4e="069 Enfances Gauvain, Meyer.docx"
nome_f5e="070 Eracle, Gautier d’Arras, de Lage.docx"
nome_f6e="071 Erec, en prose, Colombo Timelli.docx"
nome_f7e="072 Escanor, Trachsler.docx"
nome_f8e="073 Escanor, pt. II, Tracshler.docx"
nome_f9e="074 Esclarmonde, alessandrini, Schaefer.docx"
nome_f10e="075ESC~1.DOC"
nome_f11e="076 Escoufle, Sweetser.docx"
nome_f12e="077 Estoire du St. Graal, t. I, Ponceau.docx"
nome_f13e="078 Estoire du Graal, t. II, Ponceau.docx"
nome_f14e="079 Eustache le Moine, roman de, Holden.docx"

nome_f1f="Fergus INTERNET.docx"
nome_f2f="Fergus, Frescoln.docx"
nome_f3f="Floire et Blanchefleur, d_Orbigny, Leclanche.docx"
nome_f4f="Floire et Blanchefleur, Pelan.docx"
nome_f5f="Florence de Rome, II Wallenskold.docx"
nome_f6f="Florence de Rome, Wallenskold.docx"
nome_f7f="Floriant et Florete, Combes Trachsler.docx"
nome_f8f="Floriant et Florette, Levy.pdf.docx"
nome_f9f="Floris et Liryopé Barrette.docx"
nome_f10f="Flourence de Rome, Crisler.docx"
nome_f11f="Fouke Fitz Waryn, Lecco.docx"
nome_f12f="Froissart, Melyador, Bragantini-Maillard.docx"

nome_f1g="Galeran de Bretagne, Dufornet.docx"
nome_f2g="Gerard de Nevers, Histoire de, Marchal.docx"
nome_f3g="Gerbert de Montreuil, Continuation Perceval, Le Nan.docx"
nome_f4g="Gilles de Chine, Liétard-Rouzé.docx"
nome_f5g="Gillion de Trazegnies, Vincent.docx"
nome_f6g="Gliglois, roman de, Lemaire.docx"
nome_f7g="Godin, Meunier.docx"
nome_f8g="Gui de Warewic, Ewert, pt. 2.docx"
nome_f9g="Gui de Warewic, Ewert, pt.1.docx"
nome_f10g="Guillaume de Palerne, Ferlampin Acher.docx"
nome_f11g="Guiron, roman de, continuation, Veneziale.docx"
nome_f12g="Guiron, roman de, parte prima, Lagomrasini.docx"
nome_f13g="Guiron, roman de, parte seconda, Stefanelli.docx"
nome_f14g="Guy de Warwick, en prose, Conlon.docx"

nome_f1h="Histoire de Gylles de Chin, Place.docx"
nome_f2h="Histoire de Jason, Pinkernell.docx"
nome_f3h="Histoire de la reine Berthe, Tylus.docx"
nome_f4h="Horn, roman de, Pope.docx"
nome_f5h="Huon de Bordeaux, Kibler.docx"

nome_f1i="Ille et Galeron, Lefevre.docx"
nome_f2i="Ipomedon, roman de, Holden.docx"

#ogni testo viene elaborato dalla funzione getScore() che mi restituisce i nomi,i cognomi e i contesti presenti in un testo
f1=getScore("testi/B/B/Balain, roman de, Dominica Legge.docx",nome_f1)
f2=getScore("testi/B/B/Barlaam et Josaphat, Chardry.docx",nome_f2)
f3=getScore("testi/B/B/Barlaam et Josaphat, Gui de Cambrai, Appel.docx",nome_f3)
f4=getScore("testi/B/B/Barlaam et Josaphat, Mills.docx",nome_f4)
f5=getScore("testi/B/B/Barlaam et Josaphat, monte Athos, Agrigoraei.docx",nome_f5)
f6=getScore("testi/B/B/Barlaam et Josaphat, Sonet, t. I.docx",nome_f6)
f7=getScore("testi/B/B/Barlaam et Josaphat, Sonet.docx",nome_f7)
f8=getScore("testi/B/B/Baudouin de Flandre, Pinto-Mathieu.docx",nome_f8)
f9=getScore("testi/B/B/Bel Inconnu, Perret Weil.docx",nome_f9)
f10=getScore("testi/B/B/Belle Helene de Costantinople, Jean Wauquelin, Crecy 2022.docx",nome_f10)
f11=getScore("testi/B/B/Belle Helene de Costantinople, Roussel.docx",nome_f11)
f12=getScore("testi/B/B/Bérinus, prosa, Boussat.docx",nome_f12)
f13=getScore("testi/B/B/Bérinus, versi, Bossaut.docx",nome_f13)
f14=getScore("testi/B/B/Berta de li gran pié, Morgan.docx",nome_f14)
f15=getScore("testi/B/B/Berte, histoire de la reine, Tylus.docx",nome_f15)
f16=getScore("testi/B/B/Bertes aus grans piés, Adenet le Roy, Henry.docx",nome_f16)
f17=getScore("testi/B/B/Biaodouz, Lemaire.docx",nome_f17)
f18=getScore("testi/B/B/Blancandin, en prose, Greco.docx",nome_f18)
f19=getScore("testi/B/B/Brun de la Montagne.docx",nome_f19)
f20=getScore("testi/B/B/Brut, en prose.docx",nome_f20)
f21=getScore("testi/B/B/Brut, roman de, Wace, Arnold.docx",nome_f21)

f1c=getScore("testi/C/Cardenois, roman de, Cocco.docx",nome_f1c)
f2c=getScore("testi/C/Cassidorus, roman de, t. I, Palermo.docx",nome_f2c)
f3c=getScore("testi/C/Cassidorus, roman de, t. II, Palermo.docx",nome_f3c)
f4c=getScore("testi/C/Charles de Hongrie, roman de, Chenerie.docx",nome_f4c)
f5c=getScore("testi/C/Charrette, conte de, Combes.docx",nome_f5c)
f6c=getScore("testi/C/chastelain de Vergi, Stuip.docx",nome_f6c)
f7c=getScore("testi/C/Chastelaine de Coucy, Babbi.docx",nome_f7c)
f8c=getScore("testi/C/Chatelain de Coucy et dame de Fayel, roman de, Gaullier-Bougassas.docx",nome_f8c)
f9c=getScore("testi/C/Cheval de Fust, Meliacin, ed. Saly.docx",nome_f9c)
f10c=getScore("testi/C/Chevalerie de Judas Macabée, Smeets.docx",nome_f10c)
f11c=getScore("testi/C/Chevalier à l_épée, Johnston-Owen.docx",nome_f11c)
f12c=getScore("testi/C/Chevalier as deus espees, le, Rockwell.docx",nome_f12c)
f13c=getScore("testi/C/Chevalier au papegau, conte de, Charpentier-Victorin.docx",nome_f13c)
f14c=getScore("testi/C/Ciperis de Vignevaux, Ramello.docx",nome_f14c)
f15c=getScore("testi/C/Claris et Laris, Pierreville.docx",nome_f15c)
f16c=getScore("testi/C/Clarisse et Florent.docx",nome_f16c)
f17c=getScore("testi/C/Cleomades, en prose, Trachsler Mailliet.docx",nome_f17c)
f18c=getScore("testi/C/Cleomadés, Henry.docx",nome_f18c)
f19c=getScore("testi/C/Cleriadus et Meliadice, Zink.docx",nome_f19c)
f20c=getScore("testi/C/Cligés, Chretien de Troyes, Foerster.docx",nome_f20c)
f21c=getScore("testi/C/Cligés, en prose, Colombo-Timelli.docx",nome_f21c)
f22c=getScore("testi/C/Comte d_Artois, roman de, Seignereut.docx",nome_f22c)
f23c=getScore("testi/C/comte de Poitiers, roman de, Malmerg.docx",nome_f23c)
f24c=getScore("testi/C/Cristal et Clarie, Breuer.docx",nome_f24c)
f25c=getScore("testi/C/Croissant, roman de, Schafer.docx",nome_f25c)

f1d=getScore("testi/D/063 Dame au liocorne, roman de.docx",nome_f1d)
f2d=getScore("testi/D/064 Didot-Perceval, Roach.docx",nome_f2d)
f3d=getScore("testi/D/065 Durmart le Galois, Gildea INTERNET.docx",nome_f3d)

f1e=getScore("testi/E/066 Eledus et Serene, roman de, Reinall.docx",nome_f1e)
f2e=getScore("testi/E/067 Eneas, roman de, Salverda de Grave.docx",nome_f2e)
f3e=getScore("testi/E/068 Eneas, roman de, t. II, Salverda de Grave.docx",nome_f3e)
f4e=getScore("testi/E/069 Enfances Gauvain, Meyer.docx",nome_f4e)
f5e=getScore("testi/E/070 Eracle, Gautier d’Arras, de Lage.docx",nome_f5e)
f6e=getScore("testi/E/071 Erec, en prose, Colombo Timelli.docx",nome_f6e)
f7e=getScore("testi/E/072 Escanor, Trachsler.docx",nome_f7e)
f8e=getScore("testi/E/073 Escanor, pt. II, Tracshler.docx",nome_f8e)
f9e=getScore("testi/E/074 Esclarmonde, alessandrini, Schaefer.docx",nome_f9e)
f10e=getScore("testi/E/075ESC~1.DOC",nome_f10e)
f11e=getScore("testi/E/076 Escoufle, Sweetser.docx",nome_f11e)
f12e=getScore("testi/E/077 Estoire du St. Graal, t. I, Ponceau.docx",nome_f12e)
f13e=getScore("testi/E/078 Estoire du Graal, t. II, Ponceau.docx",nome_f13e)
f14e=getScore("testi/E/079 Eustache le Moine, roman de, Holden.docx",nome_f14e)

f1f=getScore("testi/F/Fergus INTERNET.docx",nome_f1f)
f2f=getScore("testi/F/Fergus, Frescoln.docx",nome_f2f)
f3f=getScore("testi/F/Floire et Blanchefleur, d_Orbigny, Leclanche.docx",nome_f3f)
f4f=getScore("testi/F/Floire et Blanchefleur, Pelan.docx",nome_f4f)
f5f=getScore("testi/F/Florence de Rome, II Wallenskold.docx",nome_f5f)
f6f=getScore("testi/F/Florence de Rome, Wallenskold.docx",nome_f6f)
f7f=getScore("testi/F/Floriant et Florete, Combes Trachsler.docx",nome_f7f)
f8f=getScore("testi/F/Floriant et Florette, Levy.pdf.docx",nome_f8f)
f9f=getScore("testi/F/Floris et Liryopé Barrette.docx",nome_f9f)
f10f=getScore("testi/F/Flourence de Rome, Crisler.docx",nome_f10f)
f11f=getScore("testi/F/Fouke Fitz Waryn, Lecco.docx",nome_f11f)
f12f=getScore("testi/F/Froissart, Melyador, Bragantini-Maillard.docx",nome_f12f)

f1g=getScore("testi/G/Galeran de Bretagne, Dufornet.docx",nome_f1g)
f2g=getScore("testi/G/Gerard de Nevers, Histoire de, Marchal.docx",nome_f2g)
f3g=getScore("testi/G/Gerbert de Montreuil, Continuation Perceval, Le Nan.docx",nome_f3g)
f4g=getScore("testi/G/Gilles de Chine, Liétard-Rouzé.docx",nome_f4g)
f5g=getScore("testi/G/Gillion de Trazegnies, Vincent.docx",nome_f5g)
f6g=getScore("testi/G/Gliglois, roman de, Lemaire.docx",nome_f6g)
f7g=getScore("testi/G/Godin, Meunier.docx",nome_f7g)
f8g=getScore("testi/G/Gui de Warewic, Ewert, pt. 2.docx",nome_f8g)
f9g=getScore("testi/G/Gui de Warewic, Ewert, pt.1.docx",nome_f9g)
f10g=getScore("testi/G/Guillaume de Palerne, Ferlampin Acher.docx",nome_f10g)
f11g=getScore("testi/G/Guiron, roman de, continuation, Veneziale.docx",nome_f11g)
f12g=getScore("testi/G/Guiron, roman de, parte prima, Lagomrasini.docx",nome_f12g)
f13g=getScore("testi/G/Guiron, roman de, parte seconda, Stefanelli.docx",nome_f13g)
f14g=getScore("testi/G/Guy de Warwick, en prose, Conlon.docx",nome_f14g)

f1h=getScore("testi/H/Histoire de Gylles de Chin, Place.docx",nome_f1h)
f2h=getScore("testi/H/Histoire de Jason, Pinkernell.docx",nome_f2h)
f3h=getScore("testi/H/Histoire de la reine Berthe, Tylus.docx",nome_f3h)
f4h=getScore("testi/H/Horn, roman de, Pope.docx",nome_f4h)
f5h=getScore("testi/H/Huon de Bordeaux, Kibler.docx",nome_f5h)

f1i=getScore("testi/I/Ille et Galeron, Lefevre.docx",nome_f1i)
f2i=getScore("testi/I/Ipomedon, roman de, Holden.docx",nome_f2i)


ftot=f1+f2+f3+f4+f5+f6+f7+f8+f9+f10+f11+f12+f13+f14+f15+f16+f17+f18+f19+f20+f21+f1c+f2c+f3c+f4c+f5c+f6c+f7c+f8c+f9c+f10c+f11c+f12c+f13c+f14c+f15c+f16c+f17c+f18c+f19c+f20c+f21c+f22c+f23c+f24c+f25c+f1d+f2d+f3d+f1e+f2e+f3e+f4e+f5e+f6e+f7e+f8e+f9e+f10e+f11e+f12e+f13e+f14e+f1f+f2f+f3f+f4f+f5f+f6f+f7f+f8f+f9f+f10f+f11f+f12f+f1g+f2g+f3g+f4g+f5g+f6g+f7g+f8g+f9g+f10g+f11g+f12g+f13g+f14g+f1h+f2h+f3h+f4h+f5h+f1i+f2i
cd=pd.DataFrame(ftot,columns=["Autore","Personaggio", "Contesto"])
cd.to_excel('fatto.xlsx')
