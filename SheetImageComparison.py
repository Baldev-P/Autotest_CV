import warnings
import matplotlib.cbook
import cv2
import numpy as np
import csv
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os.path
import matplotlib.pyplot as plt
import shutil
import imutils
import stat

# "\\se4\\Y\\AutoTestBackup\\2020-04-03.2_V2.3.1beta_x64_Release\\AllOutputFiles\\Closedbugs\\006-PERFORMANCE\\SivaData\\100\\InputToAE\\SHT_1.png"
# "\\se4\\Y\\AutoTestBackup\\20200502_3_88689_2.3.1\\AllOutputFiles\\Closedbugs\\006-PERFORMANCE\\SivaData\\100\\InputToAE\\SHT_1.png"
# report = input("Input Autotest report .csv file: ")

def handleError(func, path, exc_info):
    print('Handling Error for file ', path)
    print(exc_info)
    # Check if file access issue
    if not os.access(path, os.W_OK):
        print('Hello')
        # Try to change the permision of file
        os.chmod(path, stat.S_IWUSR)
        # call the calling function again
        func(path)


if __name__ == '__main__':
    warnings.filterwarnings("ignore", category=matplotlib.cbook.mplDeprecation)
    print("Please Input the Autotest Report file(.csv) :")
    Tk().withdraw()
    report = askopenfilename()
    line = 0
    diff = 0
    outputfile = report.split(".csv")[0] + "CV.csv"
    outputdir = os.path.dirname(report)
    try:
        with open(report, 'r') as csvfile, open(outputfile, 'w', newline='') as write_obj:
            csvreader = csv.reader(csvfile)
            csv_writer = csv.writer(write_obj)
            for row in csvreader:
                line += 1
                print("Processing test case " + str(line))
                if line == 2:
                    reference = row[3]
                elif line == 3:
                    current = row[3]
                    try:
                        autotestno=row[3].split("AutoTestBackup\\")[1]
                        outputdir += "\\CVResult_"+autotestno
                    except:
                        outputdir += "\\CVResult_"
                    #outputdir = outputdir.replace("\\AllOutputFiles", "")
                    outputdir=outputdir[0:outputdir.rfind("\\")]
                    if os.path.isdir(outputdir):
                        if os.path.isdir(outputdir + "\\WMBatchFiles"):
                            shutil.rmtree(outputdir + "\\WMBatchFiles", onerror=handleError)
                        shutil.rmtree(outputdir, ignore_errors=True)  # onerror=handleError)
                    os.mkdir(outputdir)
                    batchfiledir = "CVResult_" + autotestno
                    batchfiledir = batchfiledir[0:batchfiledir.rfind("\\")]
                    batchfile = batchfiledir + "\\WMBatchFiles"
                    os.mkdir(outputdir + "\\WMBatchFiles")
                elif line == 18:
                    row.pop()
                    row.append("Sheet Layout Difference (in kPixels)")
                elif line > 18:
                    if row[3] == "Files Different from Reference Auto Test" and row[1] == "NCD":
                        image = row[2]
                        image = image.replace(".NCD", ".png")
                        image1 = cv2.imread(reference + "\\" + image)
                        image2 = cv2.imread(current + "\\" + image)
                        difference = cv2.subtract(image1, image2)
                        result = np.any(difference)
                        pixels = np.sum(image1 != image2)
                        pixels /= 1000.0
                        # difference[np.where((difference == [0, 0, 0]).all(axis=2))] = [255, 255, 255]
                        # difference = cv2.bitwise_not(difference)
                        image = image.replace("\\", "_")
                        if result:
                            """
                            plt.figure()
                            f, axarr = plt.subplots(2, 2)
                            axarr[0, 0].imshow(image1)
                            axarr[0, 1].imshow(image2)
                            axarr[1, 0].imshow(difference)
                            """
                            plt.subplot(2, 2, 1), plt.imshow(image1), plt.title("Reference")
                            plt.subplot(2, 2, 2), plt.imshow(image2), plt.title("Current")
                            plt.subplot(2, 2, 3), plt.imshow(difference), plt.title("Difference")
                            plt.tight_layout()
                            # manager = plt.get_current_fig_manager()
                            # manager.resize(*manager.window.maxsize())
                            figure = plt.gcf()  # get current figure
                            figure.set_size_inches(32, 18)  # set figure's size manually to your full screen (32x18)
                            outputfigure = outputdir + "\\" + image
                            plt.savefig(outputfigure, bbox_inches='tight')
                            plt.close()
                            """
                            final = cv2.hconcat(image1, image2)
                            final = cv2.vconcat(final, difference)
                            cv2.imwrite(outputdir + "\\" + image, final)
                            """
                            diff += 1
                            row.append("")
                            row.append("=HYPERLINK(\"%s\",\"%.3f\")" % (batchfiledir + "\\" + image, pixels))
                    if row[3] == "Files Different from Reference Auto Test":
                        reffile = reference + "\\" + row[2]
                        curfile = current + "\\" + row[2]
                        reffolder = reffile[0:reffile.rfind("\\")]
                        curfolder = curfile[0:curfile.rfind("\\")]
                        try:
                            f1 = open(outputdir + "\\WMBatchFiles" + "\\" + row[2].replace("\\", "_") + "Folder.bat",
                                      "w")
                            f1.write("\"C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe\" \"%s\" \"%s\"" % (
                                reffolder, curfolder))
                            f1.close()
                            f2 = open(outputdir + "\\WMBatchFiles" + "\\" + row[2].replace("\\", "_") + ".bat", "w")
                            f2.write(
                                "\"C:\\Program Files (x86)\\WinMerge\\WinMergeU.exe\" \"%s\" \"%s\"" % (
                                    reffile, curfile))
                            f2.close()
                            row[1] = "=HYPERLINK(\"%s\",\"%s\")" % (
                                batchfile + "\\" + row[2].replace("\\", "_") + "Folder.bat", row[1])
                            row[2] = "=HYPERLINK(\"%s\",\"%s\")" % (
                                batchfile + "\\" + row[2].replace("\\", "_") + ".bat", row[2])
                        except:
                            print("Error linking winmerge for file comparison")

                csv_writer.writerow(row)
        csvfile.close()
        write_obj.close()
        print("Sheets having difference in Layout: %d.\nTotal files checked: %d." % (diff, line-18))
        print("Processing completed. \nCheck the output here: " + outputfile)
        input("Press any key to continue...")
    except IOError:
        print("Unexpected error occured.")
