"""extracts TNEF encoded content from for example winmail.dat attachments.
"""
import sys, logging

logger = logging.getLogger("tnef-decode")

from util import bytes_to_int, checksum
from mapi import TNEFMAPI_Attribute, decode_mapi


class TNEFObject:
   "a TNEF object that may contain a property and an attachment"

   def __init__(self, data, do_checksum=False):
      offset = 0
      self.level = bytes_to_int(data[offset: offset+1]); offset += 1
      self.name = bytes_to_int(data[offset:offset+2]); offset += 2
      self.type = bytes_to_int(data[offset:offset+2]); offset += 2
      att_length = bytes_to_int(data[offset:offset+4]); offset += 4
      self.data = data[offset:offset+att_length]; offset += att_length
      att_checksum = bytes_to_int(data[offset:offset+2]); offset += 2

      self.length = offset

      if do_checksum:
         calc_checksum = checksum(self.data)
         if calc_checksum != att_checksum:
            logger.warn("Checksum: %s != %s" % (calc_checksum, att_checksum))
      else:
         calc_checksum = att_checksum

      # whether the checksum is ok
      self.good_checksum = calc_checksum == att_checksum


   def __str__(self):
      return "<%s '%s'>" % (self.__class__.__name__, TNEF.codes[self.name])


class TNEFAttachment:
   "a TNEF attachment"

   SZMAPI_UNSPECIFIED    = 0x0000 # MAPI Unspecified
   SZMAPI_NULL           = 0x0001 # MAPI null property
   SZMAPI_SHORT          = 0x0002 # MAPI short (signed 16 bits)
   SZMAPI_INT            = 0x0003 # MAPI integer (signed 32 bits)
   SZMAPI_FLOAT          = 0x0004 # MAPI float (4 bytes)
   SZMAPI_DOUBLE         = 0x0005 # MAPI double
   SZMAPI_CURRENCY       = 0x0006 # MAPI currency (64 bits)
   SZMAPI_APPTIME        = 0x0007 # MAPI application time
   SZMAPI_ERROR          = 0x000a # MAPI error (32 bits)
   SZMAPI_BOOLEAN        = 0x000b # MAPI boolean (16 bits)
   SZMAPI_OBJECT         = 0x000d # MAPI embedded object
   SZMAPI_INT8BYTE       = 0x0014 # MAPI 8 byte signed int
   SZMAPI_STRING         = 0x001e # MAPI string
   SZMAPI_UNICODE_STRING = 0x001f # MAPI unicode-string (null terminated)
   SZMAPI_SYSTIME        = 0x0040 # MAPI time (64 bits)
   SZMAPI_CLSID          = 0x0048 # MAPI OLE GUID
   SZMAPI_BINARY         = 0x0102 # MAPI binary
   SZMAPI_BEATS_THE_HELL_OUTTA_ME = 0x0033

   codes = {
      SZMAPI_UNSPECIFIED              :  "MAPI Unspecified",
      SZMAPI_NULL                     :  "MAPI null property",
      SZMAPI_SHORT                    :  "MAPI short (signed 16 bits)",
      SZMAPI_INT                      :  "MAPI integer (signed 32 bits)",
      SZMAPI_FLOAT                    :  "MAPI float (4 bytes)",
      SZMAPI_DOUBLE                   :  "MAPI double",
      SZMAPI_CURRENCY                 :  "MAPI currency (64 bits)",
      SZMAPI_APPTIME                  :  "MAPI application time",
      SZMAPI_ERROR                    :  "MAPI error (32 bits)",
      SZMAPI_BOOLEAN                  :  "MAPI boolean (16 bits)",
      SZMAPI_OBJECT                   :  "MAPI embedded object",
      SZMAPI_INT8BYTE                 :  "MAPI 8 byte signed int",
      SZMAPI_STRING                   :  "MAPI string",
      SZMAPI_UNICODE_STRING           :  "MAPI unicode-string (null terminated)",
      #SZMAPI_PT_SYSTIME              :  "MAPI time (after 2038/01/17 22:14:07 or before 1970/01/01 00:00:00)",
      SZMAPI_SYSTIME                  :  "MAPI time (64 bits)",
      SZMAPI_CLSID                    :  "MAPI OLE GUID",
      SZMAPI_BINARY                   :  "MAPI binary",
      SZMAPI_BEATS_THE_HELL_OUTTA_ME  :  "Unknown"
   }

   def __init__(self):
      self.mapi_attrs = []

   def long_filename(self):
      atname = TNEFMAPI_Attribute.MAPI_ATTACH_LONG_FILENAME
      attr = None
      for a in self.mapi_attrs:
         if a.name == atname:
            attr = a
            break

      if attr is not None:
         fn = str(attr.data).replace('0','')
      else:
         fn = self.name

      return fn.split('\\')[-1]


   def add_attr(self, attribute):
      logger.debug("Attachment attr name: 0x%4.4x" % attribute.name)
      if attribute.name == TNEF.ATTATTACHMODIFYDATE:
         logger.debug("No date support yet!")
      elif attribute.name == TNEF.ATTATTACHMENT:
         self.mapi_attrs += decode_mapi(attribute.data)
      elif attribute.name == TNEF.ATTATTACHTITLE:
         self.name = attribute.data.strip(chr(0))  # remove any NULLs
      elif attribute.name == TNEF.ATTATTACHDATA:
         self.data = attribute.data
      else:
         logger.debug("Unknown attribute name: %s" % attribute)


   def __str__(self):
      return "<ATTCH:'%s'>" % "" #self.long_filename()


class TNEF:
   "main decoder class - start by using this"

   TNEF_SIGNATURE = 0x223e9f78
   LVL_MESSAGE = 0x01
   LVL_ATTACHMENT = 0x02

   ATTOWNER                        = 0x0000 # Owner
   ATTSENTFOR                      = 0x0001 # Sent For
   ATTDELEGATE                     = 0x0002 # Delegate
   ATTDATESTART                    = 0x0006 # Date Start
   ATTDATEEND                      = 0x0007 # Date End
   ATTAIDOWNER                     = 0x0008 # Owner Appointment ID
   ATTREQUESTRES                   = 0x0009 # Response Requested.
   ATTFROM                         = 0x8000 # From
   ATTSUBJECT                      = 0x8004 # Subject
   ATTDATESENT                     = 0x8005 # Date Sent
   ATTDATERECD                     = 0x8006 # Date Recieved
   ATTMESSAGESTATUS                = 0x8007 # Message Status
   ATTMESSAGECLASS                 = 0x8008 # Message Class
   ATTMESSAGEID                    = 0x8009 # Message ID
   ATTPARENTID                     = 0x800a # Parent ID
   ATTCONVERSATIONID               = 0x800b # Conversation ID
   ATTBODY                         = 0x800c # Body
   ATTPRIORITY                     = 0x800d # Priority
   ATTATTACHDATA                   = 0x800f # Attachment Data
   ATTATTACHTITLE                  = 0x8010 # Attachment File Name
   ATTATTACHMETAFILE               = 0x8011 # Attachment Meta File
   ATTATTACHCREATEDATE             = 0x8012 # Attachment Creation Date
   ATTATTACHMODIFYDATE             = 0x8013 # Attachment Modification Date
   ATTDATEMODIFY                   = 0x8020 # Date Modified
   ATTATTACHTRANSPORTFILENAME      = 0x9001 # Attachment Transport Filename
   ATTATTACHRENDDATA               = 0x9002 # Attachment Rendering Data
   ATTMAPIPROPS                    = 0x9003 # MAPI Properties
   ATTRECIPTABLE                   = 0x9004 # Recipients
   ATTATTACHMENT                   = 0x9005 # Attachment
   ATTTNEFVERSION                  = 0x9006 # TNEF Version
   ATTOEMCODEPAGE                  = 0x9007 # OEM Codepage
   ATTORIGNINALMESSAGECLASS        = 0x9008 # Original Message Class

   codes = {
      ATTOWNER                        :  "Owner",
      ATTSENTFOR                      :  "Sent For",
      ATTDELEGATE                     :  "Delegate",
      ATTDATESTART                    :  "Date Start",
      ATTDATEEND                      :  "Date End",
      ATTAIDOWNER                     :  "Owner Appointment ID",
      ATTREQUESTRES                   :  "Response Requested",
      ATTFROM                         :  "From",
      ATTSUBJECT                      :  "Subject",
      ATTDATESENT                     :  "Date Sent",
      ATTDATERECD                     :  "Date Received",
      ATTMESSAGESTATUS                :  "Message Status",
      ATTMESSAGECLASS                 :  "Message Class",
      ATTMESSAGEID                    :  "Message ID",
      ATTPARENTID                     :  "Parent ID",
      ATTCONVERSATIONID               :  "Conversation ID",
      ATTBODY                         :  "Body",
      ATTPRIORITY                     :  "Priority",
      ATTATTACHDATA                   :  "Attachment Data",
      ATTATTACHTITLE                  :  "Attachment File Name",
      ATTATTACHMETAFILE               :  "Attachment Meta File",
      ATTATTACHCREATEDATE             :  "Attachment Creation Date",
      ATTATTACHMODIFYDATE             :  "Attachment Modification Date",
      ATTDATEMODIFY                   :  "Date Modified",
      ATTATTACHTRANSPORTFILENAME      :  "Attachment Transport Filename",
      ATTATTACHRENDDATA               :  "Attachment Rendering Data",
      ATTMAPIPROPS                    :  "MAPI Properties",
      ATTRECIPTABLE                   :  "Recipients",
      ATTATTACHMENT                   :  "Attachment",
      ATTTNEFVERSION                  :  "TNEF Version",
      ATTOEMCODEPAGE                  :  "OEM Codepage",
      ATTORIGNINALMESSAGECLASS        :  "Original Message Class"
   }

   def __init__(self, data, do_checksum = True):
      self.signature = bytes_to_int(data[0:4])
      if self.signature != TNEF.TNEF_SIGNATURE:
         sys.exit("Wrong TNEF signature: 0x%2.8x" % self.signature)
      self.key = bytes_to_int(data[4:6])
      self.objects = []
      self.attachments = []
      self.mapiprops = []
      offset = 6

      if not do_checksum:
         logger.info("Skipping checksum for performance")

      while (offset < len(data)):
         obj = TNEFObject(data[offset: offset+len(data)], do_checksum)
         offset += obj.length
         self.objects.append(obj)

         # handle attachments
         if obj.name == TNEF.ATTATTACHRENDDATA:
            attachment = TNEFAttachment()
            self.attachments.append(attachment)
         elif obj.level == TNEF.LVL_ATTACHMENT:
            attachment.add_attr(obj)

         # handle MAPI properties
         elif obj.name == TNEF.ATTMAPIPROPS:
            self.mapiprops = decode_mapi(obj.data)

            # handle BODY property
            for p in self.mapiprops:
               if p.name == TNEFMAPI_Attribute.MAPI_BODY:
                  self.body = p.data
               elif p.name == TNEFMAPI_Attribute.MAPI_BODY_HTML:
                  self.htmlbody = p.data
         else:
            logger.warn("Unknown TNEF Object: %s" % obj)


   def has_body(self):
      return True if (self.body or self.htmlbody) else False


   def __str__(self):
      atts = (", %i attachments" % len(self.attachments)) if self.attachments else ''
      return "<%s:0x%2.2x%s>" % (self.__class__.__name__, self.key, atts)