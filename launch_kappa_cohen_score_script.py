# -*- coding: utf-8 -*-
"""
Created on Tue Sep  3 16:40:01 2019

@author: Sébastien Polvent
"""

import time
import class_kappa_cohen_score as kcs


#time marker, start of processing
TIMER = time.time()
print('Please wait while processing...\n', flush=True)

kappa_score = kcs.WeightedCohenKappaScore()

print('Importing data...')
#scorer1
print(kappa_score.scorer1_name + "'s scoring...")
kappa_score.import_scoring_from_many_files()
#scorer2
print(kappa_score.scorer2_name + "'s scoring...")
kappa_score.import_scoring_from_1_file()
print('Importation finished !\n')

print("Computing Cohen's Kappa scores...")
kappa_score.compute_kappa()
kappa_score.compute_global_kappa()
print("Cohen's Kappa scores OK.\n")

print("Computing Kripendorff Alpha scores...")
kappa_score.compute_krippendorff_alpha()
print("Kripendorff Alpha scores OK.\n")

print('Saving results to Excel...')
kappa_score.save_results_to_excel()
print('Data saved.\n')

print('Highlighting scoring disagreements in ' + kappa_score.scorer2_name + ' file...\n')
kappa_score.highlight_differences_scorer2()

print('Done !\n', flush=True)
print("Result file : " + kappa_score.result_filename)
TEMPS_INTER = time.time() - TIMER
print(f'Processing time : {TEMPS_INTER:.2f} seconds.')
