package example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Vector;

import com.pff.PSTAttachment;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

public class Test {

    static boolean saveAttachmentsToFS = false;
    static boolean verbose = false;

    public static void main(final String[] args) {
        try {
            if (saveAttachmentsToFS) {
                File tmpDir = new File("/media/sf_Outlook/0java-libpst/");
                tmpDir.mkdir();
            }

            File dir = new File("/media/sf_Outlook/test");
            File[] directoryListing = dir.listFiles();
            Arrays.sort(directoryListing);
            if (directoryListing != null) {
                for (File child : directoryListing) {
                    System.out.println(child.getPath());
                    new Test(child.getPath());
                }
            }
        } catch (final Exception err) {
            err.printStackTrace();
        }
    }

    public Test(final String filename) {
        try {
            final PSTFile pstFile = new PSTFile(filename);
            System.out.println(pstFile.getMessageStore().getDisplayName());
            long start = System.currentTimeMillis();
            this.processFolder(pstFile.getRootFolder());
            long end = System.currentTimeMillis();
            System.out.printf("processed in %d ms\n", (end - start));
        } catch (final Exception err) {
            err.printStackTrace();
        }
    }

    int depth = -1;
    int tmpDirIndex = 1;

    public void processFolder(final PSTFolder folder) throws PSTException, java.io.IOException {
        this.depth++;
        // the root folder doesn't have a display name
        if (this.depth > 0) {
            this.printDepth();
            System.out.println(folder.getDisplayName());
        }

        // go through the folders...
        if (folder.hasSubfolders()) {
            final Vector<PSTFolder> childFolders = folder.getSubFolders();
            for (final PSTFolder childFolder : childFolders) {
                this.processFolder(childFolder);
            }
        }

        // and now the emails for this folder
        if (folder.getContentCount() > 0) {
            this.depth++;
            PSTMessage email = (PSTMessage) folder.getNextChild();
            while (email != null) {
                if (verbose) {
                    this.printDepth();
                    System.out.println("Email: " + email.getDescriptorNodeId() + " - " + email.getSubject());
                } else {
                    this.printDot();
                }
                if (email.hasAttachments() && saveAttachmentsToFS) {
                    // make a temp dir for the attachments
                    File tmpDir = new File("/media/sf_Outlook/0java-libpst/" + tmpDirIndex);
                    tmpDir.mkdir();

                    // walk list of attachments and save to fs
                    for (int i = 0; i < email.getNumberOfAttachments(); i++) {
                        PSTAttachment attachment = email.getAttachment(i);
                        String filename = attachment.getFilename();

                        if (filename.trim().length() > 0) {
                            filename = "/media/sf_Outlook/0java-libpst/" + tmpDirIndex + "/" + filename;
                            if (verbose) {
                                System.out.printf("saving attachment to %s\n", filename);
                            }

                            final FileOutputStream out = new FileOutputStream(filename);
                            final InputStream attachmentStream = attachment.getFileInputStream();
                            final int bufferSize = 8176;
                            final byte[] buffer = new byte[bufferSize];
                            int count;
                            do {
                                count = attachmentStream.read(buffer);
                                out.write(buffer, 0, count);
                            } while (count == bufferSize);
                            out.close();
                        }
                    }
                    tmpDirIndex++;
                }
                email = (PSTMessage) folder.getNextChild();
            }
            this.depth--;
        }
        this.depth--;
    }

    int col = 0;
    private void printDot() {
        System.out.print(".");
        if (col++ > 100) {
            System.out.println("");
            col = 0;
        }
    }


    public void printDepth() {
        if (col > 0) {
            col = 0;
            System.out.println("");
        }
        for (int x = 0; x < this.depth - 1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
}
