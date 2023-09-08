import pandas as pd
from openpyxl import load_workbook



def analyzing_questionnaire(input_file_path, output_file_path, questionnaire_file_path):
    # Read the 'Responses', 'embodiment' and 'Characteristics' sheets from the input Excel files
    responses_df = pd.read_excel(input_file_path, sheet_name='Responses')
    characteristics_df = pd.read_excel(questionnaire_file_path, sheet_name='Characteristics')
    embodiment_df = pd.read_excel(questionnaire_file_path, sheet_name='Embodiment')

    # Creating a least of all the participant's avatars
    Avatars = responses_df['Avatar'].unique()


    # Create the pivot table
    pivot_table_df = responses_df.melt(id_vars=['Avatar'], var_name='Question', value_name='Responses')
    pivot_table_df = pivot_table_df.pivot(index='Question', columns='Avatar', values='Responses')
    pivot_table_df.reset_index(inplace=True)
    pivot_table_df['Question'] = pivot_table_df['Question'].str.strip()  # Remove any extra spaces

    # Perform the join operation for the self-esteem/ Proteus effect statements.
    proteus_df = pd.merge(characteristics_df, pivot_table_df, left_on='Question', right_on='Question', how='left')
    proteus_df.drop('Question', axis=1, inplace=True)  # Drop the redundant column
    def apply_reverse_scoring(row):
        if row['reverse scored'] == 'reverse scored':
            if row['attribute'] == 'Self esteem':
                return 5 - row[Avatars]
            else:
                return 8 - row[Avatars]
        else:
            return row[Avatars]

    # Apply the reverse scoring
    proteus_df[Avatars] = proteus_df.apply(apply_reverse_scoring, axis=1)
    # Create an empty DataFrame to store the resulting matrix of the proteus effect and
    corrected_matrix_df = pd.DataFrame(index=Avatars)

    # Calculate the sum of scores for each unique value in the 'Avatar' column
    for avatar in proteus_df['Avatar / Rosenberg'].unique():
        avatar_df = proteus_df[proteus_df['Avatar / Rosenberg'] == avatar]
        avatar_scores = avatar_df[Avatars].sum()
        corrected_matrix_df[avatar] = avatar_scores

    def enmbodiment_score():

        # Strip extra whitespaces and newline characters from the "Question" column in the embodiment table
        embodiment_df['Question'] = embodiment_df['Question'].str.strip()
        # Perform the join operation for the embodiment effect.
        embodiment_score_df = pd.merge(embodiment_df, pivot_table_df, left_on='Question', right_on='Question', how='left')
        #embodiment_score_df.drop('Question', axis=1, inplace=True)  # Drop the redundant column
        def embodiment_score_adaptation(row):
            # Perform adaptation to the scoring
            row[Avatars] = row[Avatars].replace({1: -3, 2: -2, 3: -1, 4: 0, 5: 1, 6: 2, 7: 3})
            return row[Avatars]

        # Apply the embodiment_score_adaptation function
        embodiment_score_df[Avatars] = embodiment_score_df.apply(embodiment_score_adaptation, axis=1)


    # calculating the embodiment score in different criteria following the embodiment standarize questionare: https://www.frontiersin.org/articles/10.3389/frobt.2018.00074/full#note2
        embodiment_components = {
            'Ownership': [],
            'Agency': [],
            'Tactile Sensations': [],
            'Location': [],
            'Appearance': [],
            'Total Embodiment': []
        }

        for avatar in Avatars:
            Ownership = (embodiment_score_df.loc[0, avatar] - embodiment_score_df.loc[1, avatar]) - embodiment_score_df.loc[2, avatar] + (embodiment_score_df.loc[3, avatar] - embodiment_score_df.loc[4, avatar])
            Agency = embodiment_score_df.loc[5, avatar] + embodiment_score_df.loc[6, avatar] + embodiment_score_df.loc[7, avatar] - embodiment_score_df.loc[8, avatar]
            Tactile_Sensations = (embodiment_score_df.loc[9, avatar] - embodiment_score_df.loc[10, avatar]) + embodiment_score_df.loc[11, avatar] + embodiment_score_df.loc[12, avatar]
            Location = embodiment_score_df.loc[13, avatar] - embodiment_score_df.loc[14, avatar]
            Appearance = embodiment_score_df.loc[15, avatar] + embodiment_score_df.loc[16, avatar] + embodiment_score_df.loc[17, avatar] + embodiment_score_df.loc[18, avatar]
            Total_Embodiment = ((Ownership / 5) * 2 + (Agency / 4) * 2 + Tactile_Sensations / 4 + (Location / 3) * 2 + Appearance / 4) / 8

            embodiment_components['Ownership'].append(Ownership)
            embodiment_components['Agency'].append(Agency)
            embodiment_components['Tactile Sensations'].append(Tactile_Sensations)
            embodiment_components['Location'].append(Location)
            embodiment_components['Appearance'].append(Appearance)
            embodiment_components['Total Embodiment'].append(Total_Embodiment)
            #print(embodiment_components)

        embodiment_matrix = pd.DataFrame(embodiment_components, index=Avatars)

        return embodiment_matrix

    #Checking if the questionaire has embodiment responses
    if pivot_table_df.shape[0]>30:

        # Save the joined and updated DataFrame, and the corrected matrix to the Excel sheet
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            proteus_df.to_excel(writer, sheet_name='Proteus responses', index=False)
            corrected_matrix_df.to_excel(writer, sheet_name='Self-esteem score', index=True)
            enmbodiment_score().to_excel(writer, sheet_name='Embodiment responses', index=True)



    else:
        # Save the joined and updated DataFrame, and the corrected matrix to the Excel sheet
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            proteus_df.to_excel(writer, sheet_name='Self-esteem score', index=False)
            corrected_matrix_df.to_excel(writer, sheet_name='Corrected Matrix', index=True)

    print(f"Analyzing tables have been saved to {output_file_path}")

# Example usage
questionnaire_file_path = 'questionnaire.xlsx'
input_file_path_questionnaire_1 = 'pre_experiment_5_9.xlsx'
input_file_path_questionnaire_2 = 'post_experiment_5_9.xlsx'
output_file_path_post = 'post_analyze_5_9.xlsx'
output_file_path_pre = 'pre_analyze_5_9.xlsx'
analyzing_questionnaire(input_file_path_questionnaire_2, output_file_path_post, questionnaire_file_path)
analyzing_questionnaire(input_file_path_questionnaire_1, output_file_path_pre, questionnaire_file_path)