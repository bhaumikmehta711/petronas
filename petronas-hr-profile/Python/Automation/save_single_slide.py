import os
import re
import config as conf

class save_single_slide:
    """
    docstring
    """

    def __init__(self, ppt_instance, pptx_file_path, slide_index, save_path, filename, thmx_file_path):
        self.pptx_file_path = pptx_file_path
        self.slide_index = slide_index
        self.save_path = save_path
        self.filename = filename
        self.ppt_instance = ppt_instance
        self.thmx_file_path = thmx_file_path

    def save_slide(self):
        # open the powerpoint presentation headless in background
        read_only = True
        has_title = False
        window = False
        
        prs = self.ppt_instance.Presentations.Open(self.pptx_file_path, read_only, has_title, window)
        prs2 = self.ppt_instance.Presentations.Add(WithWindow=False)
        # prs2.ApplyTemplate(os.path.abspath(r"data_migration\PPT Theme\PETRONAS.thmx"))
        prs2.ApplyTemplate(os.path.abspath(self.thmx_file_path))

        nr_slide = self.slide_index
        # insert_index = 1
        for insert_index, slide in enumerate(nr_slide, start=1):
            prs.Slides(slide).Copy()
            prs2.Slides.Paste(Index=insert_index)

        # prs2.ApplyTemplate(os.path.abspath(r"data_migration\PPT Theme\PETRONAS.thmx"))
        prs2.ApplyTemplate(os.path.abspath(self.thmx_file_path))
        # job_blob_files_dir = '../data_migration/Job_SPUR/BlobFiles/'
        save_path_file = os.path.abspath(self.save_path + "\\\\" + self.filename)
        save_path_file = re.sub(r"pptx$", r"pdf", save_path_file)
        
        formatType = 32

        os.makedirs(os.path.dirname(save_path_file), exist_ok=True)

        prs2.SaveAs(save_path_file, formatType)
        prs.Close()
        prs2.Close()

        return save_path_file

        # kills ppt_instance
        # ppt_instance.Quit()
        # del ppt_instance