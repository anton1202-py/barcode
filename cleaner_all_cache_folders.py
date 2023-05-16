
import glob
import os
import shutil




def cleaner_cache_folders():
    folder = 'cache_dir_3/'
    for filename in glob.glob(os.path.join(folder, "*")):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(filename) or os.path.islink(filename):
                os.unlink(filename)
            elif os.path.isdir(filename):
                shutil.rmtree(filename)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (filename, e))
    
    