import streamlit as str
import pandas as pd
import os

# Όνομα του αρχείου Excel για την αποθήκευση
EXCEL_FILE = "aias_database.xlsx"

# Συνάρτηση για φόρτωση ή δημιουργία του Excel
def load_data():
    if not os.path.exists(EXCEL_FILE):
        # Αν δεν υπάρχει, φτιάχνουμε τις καρτέλες
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_athletes = pd.DataFrame(columns=["ID", "Όνομα", "Έτος Γέννησης", "Θέση"])
            df_stats = pd.DataFrame(columns=["ID_Αθλητή", "Ταχύτητα 20m (s)", "Agility T-Test (s)", "Κατακόρυφο Άλμα (cm)"])
            df_athletes.to_excel(writer, sheet_name="Athletes", index=False)
            df_stats.to_excel(writer, sheet_name="Stats", index=False)
    
    athletes = pd.read_excel(EXCEL_FILE, sheet_name="Athletes")
    stats = pd.read_excel(EXCEL_FILE, sheet_name="Stats")
    return athletes, stats

def save_data(athletes_df, stats_df):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        athletes_df.to_excel(writer, sheet_name="Athletes", index=False)
        stats_df.to_excel(writer, sheet_name="Stats", index=False)

# Ρύθμιση Σελίδας
str.set_page_config(page_title="Aias Sports Club", layout="centered")
str.title("🏀 AIAS SPORTS CLUB - MANAGEMENT SYSTEM")

# Φόρτωση δεδομένων
athletes, stats = load_data()

# Μενού Πλοήγησης στην κορυφή
menu = str.sidebar.radio("Μενού Επιλογών", ["Λίστα & Μετρήσεις", "Καταχώρηση Νέου Αθλητή"])

if menu == "Καταχώρηση Νέου Αθλητή":
    str.header("📝 Φόρμα Εγγραφής Αθλητή")
    
    with str.form("register_form", clear_on_submit=True):
        name = str.text_input("Ονοματεπώνυμο Αθλητή:")
        year = str.number_input("Έτος Γέννησης:", min_value=1950, max_value=2026, value=2012, step=1)
        position = str.text_input("Θέση / Ειδικότητα:")
        
        submitted = str.form_submit_button("Αποθήκευση Αθλητή")
        if submitted:
            if name.strip() == "":
                str.error("Το όνομα είναι υποχρεοντικό!")
            else:
                next_id = int(athletes["ID"].max() + 1) if not athletes.empty else 1
                new_athlete = pd.DataFrame([{"ID": next_id, "Όνομα": name.strip(), "Έτος Γέννησης": year, "Θέση": position.strip()}])
                athletes = pd.concat([athletes, new_athlete], ignore_index=True)
                save_data(athletes, stats)
                str.success(f"Ο αθλητής {name} καταχωρήθηκε επιτυχώς!")

elif menu == "Λίστα & Μετρήσεις":
    str.header("📊 Μητρώο Αθλητών & Αθλητικές Μετρήσεις")
    
    if athletes.empty:
        str.warning("Δεν υπάρχουν καταχωρημένοι αθλητές ακόμα.")
    else:
        # Ένωση των δύο πινάκων για εμφάνιση
        merged_df = pd.merge(athletes, stats, left_on="ID", right_on="ID_Αθλητή", how="left")
        merged_df = merged_df.drop(columns=["ID_Αθλητή"]).fillna("-")
        
        str.dataframe(merged_df, use_container_width=True)
        
        str.markdown("---")
        str.subheader("⏱️ Προσθήκη Νέας Μέτρησης (Combine Drills)")
        
        # Επιλογή αθλητή για προσθήκη stats
        athlete_options = {row["Όνομα"]: row["ID"] for _, row in athletes.iterrows()}
        selected_athlete_name = str.selectbox("Επίλεξε Αθλητή:", list(athlete_options.keys()))
        selected_id = athlete_options[selected_athlete_name]
        
        col1, col2, col3 = str.columns(3)
        with col1:
            speed = str.text_input("Ταχύτητα 20m (s):", placeholder="π.χ. 3.20")
        with col2:
            agility = str.text_input("Agility T-Test (s):", placeholder="π.χ. 11.50")
        with col3:
            jump = str.text_input("Άλμα (cm):", placeholder="π.χ. 45")
            
        if str.button("Αποθήκευση Μετρήσεων"):
            if not (speed or agility or jump):
                str.error("Πρέπει να συμπληρώσεις τουλάχιστον μία μέτρηση!")
            else:
                # Αφαίρεση παλιάς μέτρησης αν υπάρχει για να έχουμε την πιο πρόσφατη
                stats = stats[stats["ID_Αθλητή"] != selected_id]
                
                new_stat = pd.DataFrame([{
                    "ID_Αθλητή": selected_id,
                    "Ταχύτητα 20m (s)": speed if speed else "-",
                    "Agility T-Test (s)": agility if agility else "-",
                    "Κατακόρυφο Άλμα (cm)": jump if jump else "-"
                }])
                stats = pd.concat([stats, new_stat], ignore_index=True)
                save_data(athletes, stats)
                str.success(f"Οι μετρήσεις για τον αθλητή {selected_athlete_name} ενημερώθηκαν!")
                str.rerun()

# Κουμπί για να κατεβάζεις το Excel (εμφανίζεται στο πλάι)
str.sidebar.markdown("---")
with open(EXCEL_FILE, "rb") as file:
    str.sidebar.download_button(
        label="📥 Κατέβασμα Βάσης (Excel)",
        data=file,
        file_name="aias_basketball_camp_stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
