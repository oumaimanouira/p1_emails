import imapclient
import email
from email.header import decode_header
import spacy
from datetime import datetime, timedelta
import os
import re
from dotenv import load_dotenv
from dateutil import parser
from typing import Dict, List

# Chargement des variables (file .env)
load_dotenv()

class MedicalEmailAgent:
    def __init__(self):
        """Agent intelligent de traitement des emails"""
        print("### Initialisation de l'agent IA médical ###")
        self.nlp = spacy.load("fr_core_news_sm")
        self.conn = None
        self._setup_nlp_pipeline()
        self.location_blacklist = {"bonjour", "merci", "service", "rh", "cordialement"}

    def _setup_nlp_pipeline(self):
        """Configuration du pipeline NLP """
        ruler = self.nlp.add_pipe("entity_ruler", before="ner")
        patterns = [
            {"label": "MEDICAL_PROFESSION", "pattern": [{"LOWER": {"IN": ["infirmier", "infirmière", "inf", "pab", "auxiliaire"]}}]},
            {"label": "SHIFT_TIME", "pattern": [{"LOWER": "quart"}, {"LOWER": "de"}, {"LOWER": {"IN": ["jour", "soir", "nuit"]}}]},
            {"label": "URGENCY", "pattern": [{"LEMMA": {"IN": ["urgent", "immédiat", "urgence", "asap"]}}]}
        ]
        ruler.add_patterns(patterns)

    def connect(self):
        """Connexion au serveur IMAP"""
        print("[INFO] Connexion à Gmail IMAP en cours...")
        try:
            self.conn = imapclient.IMAPClient(os.getenv("IMAP_SERVER"), ssl=True)
            self.conn.login(os.getenv("EMAIL_USER"), os.getenv("EMAIL_PASSWORD"))
            self.conn.select_folder("INBOX")
            print("[SUCCESS] Connexion réussie")
        except Exception as e:
            print(f"[ERROR] Erreur de connexion: {e}")

    def parse_email(self, raw_email):
        """Analyse les informations importantes d'un email"""
        msg = email.message_from_bytes(raw_email[b'BODY[]'])

        # Traitement du sujet
        subject = decode_header(msg["Subject"])[0][0]
        if isinstance(subject, bytes):
            subject = subject.decode(errors="ignore")

        # Traitement corps du message
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        return {
            "subject": subject,
            "body": body,
            "from": msg["From"],
            "date": msg["Date"]
        }

    def extract_requirements(self, text):
        """Analyse sémantique avec NLP et regex"""
        print("[INFO] Extraction des besoins")
        doc = self.nlp(text)

        requirements = {
            "profession": [],
            "shifts": [],
            "locations": [],
            "dates": [],
            "shift_duration": "",
            "urgence": False
        }

        # Detection NLP avec filtrage des erreurs
        for ent in doc.ents:
            if ent.label_ in ["GPE", "LOC"] and ent.text.lower() not in self.location_blacklist:
                requirements["locations"].append(ent.text)

            elif ent.label_ == "DATE":
                self._parse_date(ent.text, requirements)

            elif ent.label_ == "MEDICAL_PROFESSION":
                requirements["profession"].append(ent.text.upper())

            elif ent.label_ == "SHIFT_TIME":
                requirements["shifts"].append(ent.text.lower())

            elif ent.label_ == "URGENCY":
                requirements["urgence"] = True

        # Detection des dates avec regex en fallback
        date_matches = re.findall(r"\b(\d{2}/\d{2}/\d{4})\b", text)
        for date in date_matches:
            if date not in requirements["dates"]:
                requirements["dates"].append(date)

        # Detection de l'urgence contextuelle
        urgency_keywords = {"urgent", "urgence", "immédiat", "asap", "rapidement"}
        if any(word in text.lower() for word in urgency_keywords):
            requirements["urgence"] = True

        # Detection de la durée du shift
        requirements["shift_duration"] = self._calculate_shift_duration(text)
        
        # Nettoyage des résultats (suppression des doublons)
        requirements["profession"] = list(set(requirements["profession"]))
        requirements["shifts"] = list(set(requirements["shifts"]))
        requirements["locations"] = list(set(requirements["locations"]))

        print(f"[INFO] Resumé des besoins extraits : {requirements}")
        return requirements

    def _parse_date(self, text: str, requirements: Dict):
        """Analyse des dates"""
        try:
            date = parser.parse(text, dayfirst=True, fuzzy=True)
            if 2020 < date.year < 2030:
                requirements["dates"].append(date.strftime("%d/%m/%Y"))
        except:
            pass

    def _calculate_shift_duration(self, text: str) -> str:
        """Calcul intelligent de la durée du shift"""
        time_match = re.search(
            r"(\d{1,2}h\d{0,2})\s*(?:-|à|au)\s*(\d{1,2}h\d{0,2})", 
            text, 
            re.IGNORECASE
        )
        
        if time_match:
            try:
                start = datetime.strptime(time_match.group(1).replace("h", ":"), "%H:%M")
                end = datetime.strptime(time_match.group(2).replace("h", ":"), "%H:%M")

                if end < start:  # Gestion des shifts nocturnes
                    end += timedelta(days=1)
                
                duration = end - start
                return f"{duration.seconds // 3600}h"
            except Exception as e:
                print(f"[WARN] Erreur de calcul de durée corrigée: {e}")
        
        return ""

    def mark_as_processed(self, email_id):
        """Deplacer l'email dans le dossier 'Processed'"""
        try:
            # Verifier si le dossier "Processed" existe
            folders = [folder[-1] for folder in self.conn.list_folders()]
            if "Processed" not in folders:
                self.conn.create_folder("Processed")
            
            # Deplacer l'email
            self.conn.move([email_id], "Processed")
            print(f"[SUCCESS] Email {email_id} deplacé vers 'Processed'")
        except Exception as e:
            print(f"[ERROR] Erreur lors du déplacement de l'email {email_id} : {e}")

    def process_emails(self):
        """Récupere et analyse les emails non lus"""
        self.connect()
        try:
            emails = self.conn.search(["UNSEEN"])
            print(f"[INFO] {len(emails)} nouveaux emails non lus trouvés")

            results = []
            for email_id in emails:
                raw_email = self.conn.fetch([email_id], ["BODY[]"])[email_id]
                parsed_email = self.parse_email(raw_email)
                print(f"\n[INFO] Traitement de l'email de {parsed_email['from']}...")

                requirements = self.extract_requirements(f"{parsed_email['subject']}\n{parsed_email['body']}")
                classification = requirements["profession"][0] if requirements["profession"] else "non_classe"

                results.append({
                    "id": email_id,
                    "classification": classification,
                    "requirements": requirements,
                    "from": parsed_email["from"],
                    "date": parsed_email["date"],
                    "processed": True
                })

                self.mark_as_processed(email_id)

            return results
        except Exception as e:
            print(f"[ERROR] Erreur lors du traitement des emails : {e}")
            return []

if __name__ == "__main__":
    agent = MedicalEmailAgent()
    agent.process_emails()
