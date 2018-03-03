package example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Vector;

import com.pff.PSTAttachment;
import com.pff.PSTException;
import com.pff.PSTFile;
import com.pff.PSTFolder;
import com.pff.PSTMessage;

public class Test {
    public static void main(final String[] args) {
        try {
            File tmpDir = new File("/media/sf_Outlook/0java-libpst/");
            tmpDir.mkdir();
        } catch (final Exception err) {
            err.printStackTrace();
        }

        new Test("/media/sf_Outlook/done/2005-02.pst");
        // new Test(args[0]);
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
                this.printDepth();
                System.out.println("Email: " + email.getDescriptorNodeId() + " - " + email.getSubject());
                if (email.hasAttachments()) {
                    // make a temp dir for the attachments
                    File tmpDir = new File("/media/sf_Outlook/0java-libpst/" + tmpDirIndex);
                    tmpDir.mkdir();

                    // walk list of attachments and save to fs
                    for (int i = 0; i < email.getNumberOfAttachments(); i++) {
                        PSTAttachment attachment = email.getAttachment(i);
                        String filename = attachment.getFilename();

                        if (filename.trim() != "") {
                            filename = "/media/sf_Outlook/0java-libpst/" + tmpDirIndex + "/" + filename;
                            System.out.printf("saving attachment to %s\n", filename);

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

    public void printDepth() {
        for (int x = 0; x < this.depth - 1; x++) {
            System.out.print(" | ");
        }
        System.out.print(" |- ");
    }
}
