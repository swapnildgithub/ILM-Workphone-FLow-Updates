using System;
using System.IO;
using System.Reflection;
using Microsoft.MetadirectoryServices;
using Microsoft.MetadirectoryServices.Logging;
using System.Xml;
using System.Diagnostics;

namespace Mms_Metaverse
{
    /// <summary>
    /// Summary description for MVExtensionObject.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        IMVSynchronization[] myMVDlls;
        string PREFIX = "MVExtension";

        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize()
        {

            // Get list of files that start with PREFIX and have a .dll extension
            string[] fileNames =
                Directory.GetFiles(
                Utils.ExtensionsDirectory,
                PREFIX + "*.dll");

            int numFiles = fileNames.Length;

            // Create the array the size of the number of files we found.
            myMVDlls = new IMVSynchronization[numFiles];

            string fileName;
            Assembly assem;
            Type[] types;
            Object objLoaded = null;

            // Load the extension files into the array.
            for (int i = 0; i < numFiles; i++)
            {
                // Load the assembly.
                fileName = fileNames[i];
                assem = Assembly.LoadFrom(fileName);

                types = assem.GetExportedTypes();

                // Examine all the object types in the assembly for
                // ones that starts with "MVExtension". We assume
                // this object type is an MV extension object type.
                foreach (Type type in types)
                {
                    if (type.GetInterface("Microsoft.MetadirectoryServices.IMVSynchronization") != null)
                    {
                        // Create an instance of the MV extension object type.
                        objLoaded = assem.CreateInstance(type.FullName);
                        break;
                    }
                }

                // If an object type starting with "MVExtension" could not be found,
                // or if an instance could not be created, throw an exception.
                if (null == objLoaded)
                {
                    throw new UnexpectedDataException("Found MV extension " +
                        "DLL (" + fileName + ") without an " + PREFIX + "* type");
                }

                // Add this MV extension object to our array of objects.
                myMVDlls[i] = (IMVSynchronization)objLoaded;

                // Call the Initialize() method on each MV extension object.
                myMVDlls[i].Initialize();
            }
        }

        void IMVSynchronization.Terminate()
        {
            // Call the Terminate() method on each MV extension object.
            foreach (IMVSynchronization mvextension in myMVDlls)
            {
                mvextension.Terminate();
            }
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            // Call the Provision() method on each MV extension object.
            foreach (IMVSynchronization mvextension in myMVDlls)
            {
                mvextension.Provision(mventry);
            }
        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            bool shouldDelete = false;

            // Call the ShouldDeleteFromMV() method on each MV extension object.
            foreach (IMVSynchronization mvextension in myMVDlls)
            {
                if (mvextension.ShouldDeleteFromMV(csentry, mventry))
                {
                    shouldDelete = true;
                    break;
                }
            }

            return (shouldDelete);
        }
    }
}