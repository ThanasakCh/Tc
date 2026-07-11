import sys
import os
sys.path.append(os.path.abspath('src'))

from unittest.mock import patch
import pipeline

inputs = ['2', 'ป8401_สทล01_มส_06 EC', 'y', 'y', 'y']
def mock_input(prompt):
    print(prompt, end='')
    return inputs.pop(0) if inputs else 'n'

if __name__ == '__main__':
    with patch('builtins.input', mock_input):
        pipeline.run_pipeline()
