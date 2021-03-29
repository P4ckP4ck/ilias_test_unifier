import pandas as pd
import numpy as np
import os
from tqdm import tqdm

class IliasParser:

    def __init__(self, filename):
        non_question_cols = 19  # Anzahl der Spalten in der Ausgabedatei ohne Fragen!
        self.df_dict = pd.read_excel(filename, sheet_name=None)
        self.nr_questions = len(self.df_dict["Testergebnisse"].keys()) - non_question_cols

    @property
    def test_results(self):
        # Im folgenden werde fehlende Werte ergänzt, dafür werden unterschiedliche Methoden je nach Spalte benötigt!
        df_full = self.df_dict["Testergebnisse"].fillna(value=0)
        df_statistics = self.df_dict["Testergebnisse"].loc[:, :"Durchlauf"].fillna(method="ffill")
        df_full.update(df_statistics, overwrite=True)
        return df_full.set_index("Name")

    def _get_correct_entry(self, df, name):
        # Kleine Hilfsfunktion da teilweise mehr als eine Zeile pro Person verarbeitet werden muss
        final_rating_row = df.loc[name]["Bewerteter Durchlauf"]
        if type(final_rating_row) is pd.Series:
            boolean_comprehension = df.loc[name]["Durchlauf"] == final_rating_row[0]
            return df.loc[name][boolean_comprehension].iloc[0]
        else:
            return df.loc[name]

    def _unique_test_results(self):
        # Die Funktion wählt die korrekte Spalte im Übersichtstabellenblatt aus.
        # ILIAS erzeugt Leerzeilen, wenn eine Person mehrere Testdurchläufe durchführt
        unique_df = {}
        df = self.test_results
        for name in df.index:
            if name in unique_df:
                continue
            unique_df[name] = self._get_correct_entry(df, name)
        return pd.DataFrame(unique_df).T

    def _answers_per_sheet(self):
        answer_sheets = list(self.df_dict.keys())[1:]
        for i, name in enumerate(answer_sheets):
            df = self.df_dict[name]
            df.columns = ["Question", "Answer"]
            df = df.set_index("Question")
            df.to_csv(f"./answer_sheets/{i}.csv")

    def _answers_single_sheet(self):
        df = self.df_dict["Auswertung für alle Benutzer"]
        df.columns = ["Question", "Answer"]
        user = {}
        j = 0
        for i, line in tqdm(df.iterrows()):
            user[i] = line
            if type(line["Question"]) is str:
                if "Ergebnisse von Testdurchlauf" in line["Question"]:
                    user = pd.DataFrame(user).T.iloc[:-1]
                    user.reset_index(drop=True, inplace=True)
                    user.to_csv(f"./answer_sheets/{j}.csv")
                    j += 1
                    user = {}

    def _create_answer_log(self):
        if not "answer_sheets" in os.listdir():
            os.makedirs("./answer_sheets/")
        if "Auswertung für alle Benutzer" in self.df_dict.keys():
            self._answers_single_sheet()
        else:
            self._answers_per_sheet()

    def _create_results_dict(self):
        dir = "./answer_sheets/"
        results = os.listdir(dir)
        result_dict = ResultDict()
        unique_id = 0
        for (student_id, file) in tqdm(enumerate(results)):
            table = pd.read_csv(dir + file, index_col=0)
            for row in table.iterrows():
                if row[1].Question is np.nan and row[1].Answer is np.nan:
                    # deletes empty rows from file and skips loop execution
                    continue
                if row[1].Question in ("Formelfrage", "Single Choice", "Multiple Choice"):
                    # identifies current question
                    current_question = row[1].Answer
                    unique_id = 0  # the unique id helps, if ilias is not returning any variables as question name
                    continue
                result_dict.append(current_question, row[1], unique_id, student_id)
                unique_id += 1
        result_dict.save()

    def export(self, name):
        df = self._unique_test_results()
        self._create_answer_log()
        df.to_csv(f"{name}.csv")
        print(f"Test results saved as {name}.csv!")

    def export_anon(self, name):
        df = self._unique_test_results()
        self._create_answer_log()
        df.reset_index(drop=True, inplace=True)
        df["Benutzername"] = range(len(df))
        df["Matrikelnummer"] = range(len(df))
        df.to_csv(f"{name}.csv")
        print(f"Anonymous test results saved as {name}.csv!")


class Mat2Name:

    def __init__(self):
        self.df_list = []
        self.prepared = False
        self.lookup = pd.DataFrame()

    def append(self, df):
        # Fügt dem Helfer die aktuellste Liste der Matrikelnummern/Benutzernamen hinzu
        self.df_list.append(df)
        self.prepared = False

    def lookup_name(self, mat):
        #Gibt die angegebene Matrikelnummer als Benutzername zurück
        if not self.prepared:
            self.finish_df()
        return self.lookup.loc[mat]["Name"]

    def finish_df(self):
        # Sorgt dafür, dass die Matrikelnummer suchbar wird
        self.lookup = self.df_list[0].append(self.df_list[1:]).drop_duplicates()
        self.prepared = True


def summarize_tests():
    # Initialisierung
    dir = "./place_tests_here/"
    files = os.listdir(dir)
    results = []
    overview = []
    # Der Helfer wird benötigt um im Anschluss den Matrikelnummer die Benutzernamen zuordnen zu können
    mat_handler = Mat2Name()
    for file in tqdm(files):
        tqdm.write(f"Analysing: {file}")

        df = pd.read_excel(dir + file).fillna(method="ffill")
        mat_handler.append(df[["Matrikelnummer", "Name"]].drop_duplicates().set_index("Matrikelnummer"))

        # Die folgenden Zeilen löschen alle nicht bewerteten Durchläufe aus dem DataFrame
        unique = df.groupby("Benutzername").apply(lambda x: x[x["Durchlauf"] == x["Bewerteter Durchlauf"]])
        unique.set_index("Matrikelnummer", inplace=True)

        # Hier werden die Ergebnisse den verschiedenen Listenhelfern übergeben
        results.append(unique.loc[:, "Durchlauf":].drop("Durchlauf", axis=1).add_prefix(file + "_"))
        points_sum = pd.DataFrame(unique.loc[:, "Testergebnis in Punkten"])
        points_sum.columns = [file]  # Zur späteren Einordnung in welchem Test die Punktzahl erreicht wurde
        overview.append(points_sum)
    sh = results[0].join(results[1:])  # Wird noch nicht genutzt, ist für die IRT-Berechnung im Nachgang wichtig
    ov = overview[0].join(overview[1:])

    #Hier werden die Namen wieder den Matrikelnummern zugeordnet
    names = [mat_handler.lookup_name(mat) for mat in ov.index]
    sums = ov.sum(axis=1)
    ov.insert(0, "Summen", sums)
    ov.insert(0, "Name", names)
    ov.to_excel("Ausgabe.xlsx")

if __name__=="__main__":
    summarize_tests()