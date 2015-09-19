// Author : Matt Thorne
// Date   : 17-Sept-2015
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//  Unless required by applicable law or agreed to in writing, software
//  distributed under the License is distributed on an "AS IS" BASIS,
//  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//  See the License for the specific language governing permissions and
//  limitations under the License.
//


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OPCAutomation; // You need to reference - OPC DA Automation Wrapper 2.02

namespace RSLinxClient
{
    /// <summary>
    /// Client class for connecting to RSLinx OPC server</summary>
    /// <remarks>
    /// This class provides methods for accessing PLC tags via an OPC topic on RSLinx</remarks>
    public class RSLinxConnection
    {
       
        // OPC Server connection 
        private OPCAutomation.OPCServer RSLinxOPCServer = new OPCAutomation.OPCServer();
        // Collection to store named OPC groups
        private List<OPCAutomation.OPCGroup> RSLinxOPCGroups = new List<OPCAutomation.OPCGroup>();

        // Name of the server connection - May change depending on the version of Linx
        private static String connectionString = "RSLinx OPC Server";

        /// <summary>
        /// Connect to the RS Linx OPC server
        /// </summary>
        public void connect()
        {
            RSLinxOPCServer.Connect(connectionString);
        }

        /// <summary>
        /// Disconnect from the RS Linx OPC server and remove all tag groups
        /// </summary>
        public void disconnect()
        {
            RSLinxOPCServer.OPCGroups.RemoveAll();
            RSLinxOPCGroups.Clear();
            RSLinxOPCServer.Disconnect();
        }

        /// <summary>
        /// RS Linx OPC server connection status
        /// </summary>
        /// <returns>true if RS Linx OPC server is connected</returns>
        public bool isConnected()
        {
            //1 = connected
            //6 = disconnected
            return (RSLinxOPCServer.ServerState==1);
        }

        /// <summary>
        /// RS Linx OPC Server state
        /// </summary>
        /// <returns>Integer representing state of the connection</returns>
        public int serverState()
        {
            return RSLinxOPCServer.ServerState;
        }

        /// <summary>
        /// Gets an OPC group ID by it's name
        /// </summary>
        /// <param name="groupName">Group name to reference</param>
        /// <returns>The unique group identifier</returns>
        public int getGroupID(string groupName){
            int result=0;
            
            foreach (OPCAutomation.OPCGroup opcGroup in RSLinxOPCGroups)
            {
                if (opcGroup.Name.Equals(groupName)) result = RSLinxOPCGroups.IndexOf(opcGroup); 
            }

            return result;
        }

        /// <summary>
        /// Gets an OPC group name by it's identifier
        /// </summary>
        /// <param name="groupID">Group identifier to reference</param>
        /// <returns>The name of the group</returns>
        public string getGroupName(int groupID)
        {
            
            return RSLinxOPCGroups.ElementAt(groupID).Name;
        }

        /// <summary>
        /// Add an OPC group to the server
        /// </summary>
        /// <param name="groupName">Name of the group to add</param>
        public void addGroup(string groupName)
        {
            if (isConnected())
            {
                RSLinxOPCGroups.Add(RSLinxOPCServer.OPCGroups.Add(groupName));

                RSLinxOPCGroups.ElementAt(RSLinxOPCGroups.Count - 1).IsActive = true;
            }
        }

        /// <summary>
        /// Add a tag to a tag group
        /// </summary>
        /// <param name="groupID">Group ID to which tag should be addded</param>
        /// <param name="tagName">Tagname to add including topic</param>
        /// <remarks>Tag format should be - [TOPICNAME]TAGNAME
        /// </remarks>       
        public void addTag(int groupID, String tagName)
        {
            if ((isConnected())&&(RSLinxOPCGroups.Count==groupID+1))
            {
                int tagReference = RSLinxOPCGroups.ElementAt(groupID).OPCItems.Count + 1;
                RSLinxOPCGroups.ElementAt(groupID).OPCItems.AddItem(tagName, tagReference);
            }
        }

        /// <summary>
        /// Add a tag to a tag group
        /// </summary>
        /// <param name="groupName">Group name to which tag should be addded</param>
        /// <param name="tagName">Tagname to add including topic</param>
        /// <remarks>Tag format should be - [TOPICNAME]TAGNAME
        /// </remarks>    
        public void addTag(String groupName, String tagName)
        {

            foreach (OPCAutomation.OPCGroup opcGroup in RSLinxOPCGroups)
            {

                if (opcGroup.Name.Equals(groupName))
                {
                    int tagReference = opcGroup.OPCItems.Count + 1;
                    opcGroup.OPCItems.AddItem(tagName, tagReference);
                }
            }

        }

        /// <summary>
        /// Read all tag values from the specified group
        /// </summary>
        /// <param name="groupID">Group ID to read</param>
        /// <returns>Returns an array of tag values for the group</returns>
        public TagGroupValues readAll(int groupID)
        {
            
            System.Array syncErrors;
            System.Array tagErrors = null;
            System.Array syncValues = null;
            System.Array tagValues = null;
            System.Object qualities;
            System.Array tagQualities = null;
            System.Array qualitiesArray;
            System.Object syncTime;
            System.Array syncTimeArray;
            System.Array tagSyncTimes = null;
            System.Array tagNames = null;
            
            int itemCount = 0;
            if ((isConnected())&&(RSLinxOPCGroups.Count==groupID+1))
            {
                itemCount = RSLinxOPCGroups.ElementAt(groupID).OPCItems.Count;
                System.Array syncServerHandles = new int[itemCount + 1];

                


                //syncServerHandles.SetValue(RSLinxOPCGroup.OPCItems.Item(1).ServerHandle, 1);
                for (int i = 1; i <= itemCount; i++) syncServerHandles.SetValue(RSLinxOPCGroups.ElementAt(groupID).OPCItems.Item(i).ServerHandle, i);



                RSLinxOPCGroups.ElementAt(groupID).SyncRead(1, itemCount, ref syncServerHandles, out syncValues, out syncErrors, out qualities, out syncTime);
                             
            
                // Convert all the nasty OPCAutomation 1 based arrays to normal 0 based arrays
                tagValues = new string[itemCount];

                for (int i = 1; i <= itemCount; i++)
                {
                    tagValues.SetValue(syncValues.GetValue(i).ToString(), i - 1);
                }

                tagErrors = new string[itemCount];

                for (int i = 1; i <= itemCount; i++)
                {
                    tagErrors.SetValue(syncErrors.GetValue(i).ToString(), i - 1);
                }

                tagQualities = new string[itemCount];

                for (int i = 1; i <= itemCount; i++)
                {
                    qualitiesArray = (System.Array)qualities;
                    tagQualities.SetValue(qualitiesArray.GetValue(i).ToString(), i - 1);
                }

                tagSyncTimes = new string[itemCount];

                for (int i = 1; i <= itemCount; i++)
                {
                    syncTimeArray = (System.Array)syncTime;
                    tagSyncTimes.SetValue(syncTimeArray.GetValue(i).ToString(), i - 1);
                }

                tagNames = getTagNames(groupID);
            }

            return new TagGroupValues(tagErrors,tagNames,tagValues,tagQualities,tagSyncTimes);
        }

        /// <summary>
        /// Get all tag names for a specified group
        /// </summary>
        /// <param name="groupID">Group ID to get tag names from</param>
        /// <returns>Returns an array of tag names</returns>
        public System.Array getTagNames(int groupID)
        {
            System.Array tagNames = null;
            if ((isConnected()) && (RSLinxOPCGroups.Count == groupID + 1))
            {
                tagNames = new String[RSLinxOPCGroups.ElementAt(groupID).OPCItems.Count];

                for (int i = 0; i < RSLinxOPCGroups.ElementAt(groupID).OPCItems.Count; i++)
                {
                    tagNames.SetValue(RSLinxOPCGroups.ElementAt(groupID).OPCItems.Item(i + 1).ItemID.ToString(), i);
                }
            }


            return tagNames;

        }

        /// <summary>
        /// Get all tag names for a specified group
        /// </summary>
        /// <param name="group">Group name to get tag names from</param>
        /// <returns>Returns an array of tag names</returns>
        public System.Array getTagNames(string group)
        {
            return getTagNames(getGroupID(group));
        }


        /// <summary>
        /// Read all tag values from the specified group
        /// </summary>
        /// <param name="group">Group name to read</param>
        /// <returns>Returns an array of tag values for the group</returns>
        public TagGroupValues readAll(string group)
        {
            return readAll(getGroupID(group));
        }

        /// <summary>
        /// Read the value of a specified tag
        /// </summary>
        /// <param name="tag">Tag name to be read</param>
        /// <returns>A string representation of the tag value</returns>
        public string readTag(string tag)
        {

            
            int groupID=0;
            int tagIndex = 0;
            System.Object tagValue;
            System.Object tagQuality;
            System.Object tagSyncTime;
            //find the group
            foreach (OPCAutomation.OPCGroup opcGroup in RSLinxOPCGroups)
            {
                //find the tag and group index
                System.Array tagNames = getTagNames(RSLinxOPCGroups.IndexOf(opcGroup));
                for (int i = 0; i < tagNames.Length; i++)
                {
                    if (tagNames.GetValue(i).Equals(tag))
                    {
                        groupID = RSLinxOPCGroups.IndexOf(opcGroup);
                        tagIndex = i;
                        break;
                    }
                }

            }

            //read the tag
            RSLinxOPCGroups.ElementAt(groupID).OPCItems.Item(tagIndex).Read(1, out tagValue, out tagQuality, out tagSyncTime);
           
            //return the value
            return tagValue.ToString();
        }

        /// <summary>
        /// Write a value to the specified tag
        /// </summary>
        /// <param name="tag">Name of tag to modify</param>
        /// <param name="value">Value to write to tag</param>
        public void writeTag(string tag, object value)
        {
            
            //find the group
            foreach (OPCAutomation.OPCGroup opcGroup in RSLinxOPCGroups)
            {
                //find the tag
                System.Array tagNames = getTagNames(RSLinxOPCGroups.IndexOf(opcGroup));
                for (int i = 0; i < tagNames.Length; i++)
                {
                    if (tagNames.GetValue(i).Equals(tag))
                    {
                         
                        opcGroup.OPCItems.Item(i).Write(value);
                    }
                }

            }

        }

        /// <summary>
        /// Class for holding read tag group information
        /// </summary>
        public class TagGroupValues
        {
            public System.Array tagErrors;
            public System.Array tagNames;
            public System.Array tagValues;
            public System.Array tagQualities;
            public System.Array tagSyncTimes;

            public TagGroupValues(System.Array tagErrors, System.Array tagNames, System.Array tagValues, System.Array tagQualities, System.Array tagSyncTimes)
            {
                this.tagErrors = tagErrors;
                this.tagNames = tagNames;
                this.tagQualities = tagQualities;
                this.tagSyncTimes = tagSyncTimes;
                this.tagValues = tagValues;
            }

            public int getLength()
            {
                return this.tagValues.Length;
            }
        }
    }
}
