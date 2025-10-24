#include <pybind11/pybind11.h>
#include <exiv2/exiv2.hpp>
#include <string>
#include <sstream>
#include <iostream>


namespace py = pybind11;
const char *EXCEPTION_HINT = "Caught Exiv2 exception: ";
std::stringstream error_log;


void logHandler(int level, const char *msg)
{
    switch (level)
    {
    case Exiv2::LogMsg::debug:
    case Exiv2::LogMsg::info:
    case Exiv2::LogMsg::warn:
        std::cout << msg << std::endl;
        break;

    case Exiv2::LogMsg::error:
        // For unknown reasons, the exception thrown here cannot be caught by pybind11, so temporarily save it to error_log.
        // throw std::exception(msg);
        error_log << msg;
        break;

    default:
        return;
    }
}

/* The error log should be checked at the end of each operation.
   If there is a C++ error, it is converted to a Python exception. */
void check_error_log()
{
    std::string str = error_log.str();
    if(str != ""){
        error_log.clear();  // Clear it so that it can be used again
        error_log.str("");
        throw std::runtime_error(str);
    }
}

void init()
{
    Exiv2::LogMsg::setHandler(logHandler);
}

void set_log_level(int level)
{
    if (level == 0)
        Exiv2::LogMsg::setLevel(Exiv2::LogMsg::debug);
    if (level == 1)
        Exiv2::LogMsg::setLevel(Exiv2::LogMsg::info);
    if (level == 2)
        Exiv2::LogMsg::setLevel(Exiv2::LogMsg::warn);
    if (level == 3)
        Exiv2::LogMsg::setLevel(Exiv2::LogMsg::error);
    if (level == 4)
        Exiv2::LogMsg::setLevel(Exiv2::LogMsg::mute);
}

py::str version()
{
    return Exiv2::version();
}

// Ensure the current exiv2 version is equal to or greater than 0.27.4, which adds function Exiv2::enableBMFF().
#if EXIV2_TEST_VERSION(0,27,4)
bool enableBMFF(bool enable)
{
    return Exiv2::enableBMFF(enable);
}
#endif

#define read_block                                                     \
    {                                                                  \
        py::list table;                                                \
        for (; i != end; ++i)                                          \
        {                                                              \
            py::list line;                                             \
            line.append(py::bytes(i->key()));                          \
                                                                       \
            std::stringstream _value;                                  \
            _value << i->value();                                      \
            line.append(py::bytes(_value.str()));                      \
                                                                       \
            const char *typeName = i->typeName();                      \
            line.append(py::bytes((typeName ? typeName : "Unknown"))); \
            table.append(line);                                        \
        }                                                              \
        check_error_log();                                             \
        return table;                                                  \
    }

class Buffer{
public:
    char *data;
    long size;

    Buffer(const char *data_, long size_){
        size = size_;
        data = (char *)calloc(size, sizeof(char));
        if(data == NULL)
            throw std::runtime_error("Failed to allocate memory.");
        memcpy(data, data_, size);
    }

    void destroy(){
        if(data){
            free(data);
            data = NULL;
        }
    }

    py::bytes dump(){
        return py::bytes((char *)data, size);
    }
};

class Image{
public:
    Exiv2::Image::AutoPtr *img = new Exiv2::Image::AutoPtr;

    Image(const char *filename){
        *img = Exiv2::ImageFactory::open(filename);
        if (img->get() == 0)
            throw Exiv2::Error(Exiv2::kerErrorMessage, "Can not open this image.");
        (*img)->readMetadata();     // Calling readMetadata() reads all types of metadata supported by the image
        check_error_log();
    }

    Image(Buffer buffer){
        *img = Exiv2::ImageFactory::open((Exiv2::byte *)buffer.data, buffer.size);
        if (img->get() == 0)
            throw Exiv2::Error(Exiv2::kerErrorMessage, "Can not open this image.");
        (*img)->readMetadata();
        check_error_log();
    }

    void close_image()
    {
        delete img;
        check_error_log();
    }

    py::bytes get_bytes()
    {
        Exiv2::BasicIo &io = (*img)->io();
        return py::bytes((char *)io.mmap(), io.size());
    }

    std::string get_mime_type()
    {
        return (*img)->mimeType();
    }

    py::dict get_access_mode()
    {
        /* Get the access mode to various metadata.
        Reference:
        AccessMode Exiv2::Image::checkMode(MetadataId metadataId) const
        enum MetadataId { mdNone=0, mdExif=1, mdIptc=2, mdComment=4, mdXmp=8, mdIccProfile=16 };
        enum AccessMode { amNone=0, amRead=1, amWrite=2, amReadWrite=3 };
        */
        auto mode       = py::dict();
        mode["exif"]    = int((*img)->checkMode(Exiv2::mdExif));
        mode["iptc"]    = int((*img)->checkMode(Exiv2::mdIptc));
        mode["xmp"]     = int((*img)->checkMode(Exiv2::mdXmp));
        mode["comment"] = int((*img)->checkMode(Exiv2::mdComment));
        // mode["icc"]     = int((*img)->checkMode(Exiv2::mdIccProfile));   // Exiv2 will not check ICC
        return mode;
    }

    py::object read_exif()
    {
        Exiv2::ExifData &data         = (*img)->exifData();
        Exiv2::ExifData::iterator i   = data.begin();
        Exiv2::ExifData::iterator end = data.end();
        read_block;
    }

    py::object read_iptc()
    {
        Exiv2::IptcData &data         = (*img)->iptcData();
        Exiv2::IptcData::iterator i   = data.begin();
        Exiv2::IptcData::iterator end = data.end();
        read_block;
    }

    py::object read_xmp()
    {
        Exiv2::XmpData &data         = (*img)->xmpData();
        Exiv2::XmpData::iterator i   = data.begin();
        Exiv2::XmpData::iterator end = data.end();
        read_block;
    }

    py::object read_raw_xmp()
    {
        /*
        When readMetadata() is called, Exiv2 reads the raw XMP text,
        stores it in a string called XmpPacket, then parses it into an XmpData instance.
        */
        return py::bytes((*img)->xmpPacket());
    }

    py::object read_comment()
    {
        return py::bytes((*img)->comment());
    }

    py::object read_icc()
    {
        Exiv2::DataBuf *buf = (*img)->iccProfile();
        return py::bytes((char*)buf->pData_, buf->size_);
    }

    py::object read_thumbnail()
    {
        Exiv2::ExifThumb exifThumb((*img)->exifData());
        Exiv2::DataBuf buf = exifThumb.copy();
        return py::bytes((char*)buf.pData_, buf.size_);
    }

    void modify_exif(py::list table, py::str encoding)
    {
        // Create an empty container for storing data
        Exiv2::ExifData &exifData = (*img)->exifData();

        // Iterate the input table. each line contains a key and a value
        for (auto _line : table){

            // Convert _line from auto type to py::list type
            py::list line;
            for (auto field : _line)
                line.append(field);

            // Extract the fields in line
            std::string key      = py::bytes(line[0].attr("encode")(encoding));
            std::string value    = py::bytes(line[1].attr("encode")(encoding));
            std::string typeName = py::bytes(line[2].attr("encode")(encoding));

            // Locate the key
            Exiv2::ExifData::iterator key_pos = exifData.findKey(Exiv2::ExifKey(key));

            // Delete the existing key to set a new value, otherwise the key may contain multiple values.
            if (key_pos != exifData.end())
                exifData.erase(key_pos);

            if      (typeName == "_delete")
                continue;
            else if (typeName == "string")
                exifData[key] = value;
        }
        (*img)->setExifData(exifData);
        (*img)->writeMetadata();        // Save the metadata from memory to disk
        check_error_log();
    }

    void modify_iptc(py::list table, py::str encoding)
    {
        Exiv2::IptcData &iptcData = (*img)->iptcData();
        for (auto _line : table){
            py::list line;
            for (auto field : _line)
                line.append(field);
            std::string key      = py::bytes(line[0].attr("encode")(encoding));
            std::string typeName = py::bytes(line[2].attr("encode")(encoding));

            Exiv2::IptcData::iterator key_pos = iptcData.findKey(Exiv2::IptcKey(key));
            while (key_pos != iptcData.end()){  // Use the while loop because the iptc key may repeat
                iptcData.erase(key_pos);
                key_pos = iptcData.findKey(Exiv2::IptcKey(key));
            }

            if      (typeName == "_delete")
                continue;
            else if (typeName == "string")
            {
                std::string value = py::bytes(line[1].attr("encode")(encoding));
                iptcData[key] = value;
            }
            else if (typeName == "array")
            {
                Exiv2::Value::AutoPtr value = Exiv2::Value::create(Exiv2::string);
                for (auto item: line[1]){
                    std::string item_str = py::bytes(py::str(item).attr("encode")(encoding));
                    value->read(item_str);
                    iptcData.add(Exiv2::IptcKey(key), value.get());
                }
            }
        }
        (*img)->setIptcData(iptcData);
        (*img)->writeMetadata();
        check_error_log();
    }

    void modify_xmp(py::list table, py::str encoding)
    {
        Exiv2::XmpData &xmpData = (*img)->xmpData();
        for (auto _line : table){
            py::list line;
            for (auto field : _line)
                line.append(field);
            std::string key = py::bytes(line[0].attr("encode")(encoding));
            std::string typeName = py::bytes(line[2].attr("encode")(encoding));

            Exiv2::XmpData::iterator key_pos = xmpData.findKey(Exiv2::XmpKey(key));
            if (key_pos != xmpData.end())
                xmpData.erase(key_pos);

            if      (typeName == "_delete")
                continue;
            else if (typeName == "string")
            {
                std::string value = py::bytes(line[1].attr("encode")(encoding));
                xmpData[key] = value;
            }
            else if (typeName == "array")
            {
                Exiv2::Value::AutoPtr value = Exiv2::Value::create(Exiv2::xmpSeq);
                for (auto item: line[1]){
                    std::string item_str = py::bytes(py::str(item).attr("encode")(encoding));
                    value->read(item_str);
                }
                xmpData.add(Exiv2::XmpKey(key), value.get());
            }
        }
        (*img)->setXmpData(xmpData);
        (*img)->writeMetadata();
        check_error_log();
    }

    void modify_raw_xmp(py::str data, py::str encoding)
    {
        std::string data_str = py::bytes(data.attr("encode")(encoding));
        (*img)->setXmpPacket(data_str);
        (*img)->writeMetadata();
        (*img)->writeXmpFromPacket();   // Refresh the parsed XMP data in memory
        check_error_log();
    }

    void modify_comment(py::str data, py::str encoding)
    {
        std::string data_str = py::bytes(data.attr("encode")(encoding));
        (*img)->setComment(data_str);
        (*img)->writeMetadata();
        check_error_log();
    }

    void modify_icc(const char *data, long size)
    {
        Exiv2::DataBuf buf((Exiv2::byte *) data, size);
        (*img)->setIccProfile(buf);
        (*img)->writeMetadata();
        check_error_log();
    }

    void modify_thumbnail(const char *data, long size)
    {
        Exiv2::ExifThumb exifThumb((*img)->exifData());
        exifThumb.setJpegThumbnail((Exiv2::byte *) data, size);
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_exif()
    {
        (*img)->clearExifData();
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_iptc()
    {
        (*img)->clearIptcData();
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_xmp()
    {
        (*img)->clearXmpData();
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_comment()
    {
        (*img)->clearComment();
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_icc()
    {
        (*img)->clearIccProfile();
        (*img)->writeMetadata();
        check_error_log();
    }

    void clear_thumbnail()
    {
        Exiv2::ExifThumb exifThumb((*img)->exifData());
        exifThumb.erase();
        (*img)->writeMetadata();
        check_error_log();
    }

};


// Declare the API that needs to be mapped, to convert this CPP file into a Python module.
PYBIND11_MODULE(exiv2api, m)
{
    m.doc() = "Expose the API of exiv2 to Python.";
    m.def("init"         , &init);
    m.def("version"      , &version);
    m.def("registerNs"   , &Exiv2::XmpProperties::registerNs);
#if EXIV2_TEST_VERSION(0,27,4)
    m.def("enableBMFF"   , &enableBMFF);
#endif
    m.def("set_log_level", &set_log_level);
    py::class_<Buffer>(m, "Buffer")
        .def(py::init<const char *, long>())
        .def_readonly("data"      , &Buffer::data)
        .def_readonly("size"      , &Buffer::size)
        .def("destroy"            , &Buffer::destroy)
        .def("dump"               , &Buffer::dump);
    py::class_<Image>(m, "Image")
        .def(py::init<const char *>())
        .def(py::init<Buffer &>())
        .def("close_image"      , &Image::close_image)
        .def("get_bytes"        , &Image::get_bytes)
        .def("get_mime_type"    , &Image::get_mime_type)
        .def("get_access_mode"  , &Image::get_access_mode)
        .def("read_exif"        , &Image::read_exif)
        .def("read_iptc"        , &Image::read_iptc)
        .def("read_xmp"         , &Image::read_xmp)
        .def("read_raw_xmp"     , &Image::read_raw_xmp)
        .def("read_comment"     , &Image::read_comment)
        .def("read_icc"         , &Image::read_icc)
        .def("read_thumbnail"   , &Image::read_thumbnail)
        .def("modify_exif"      , &Image::modify_exif)
        .def("modify_iptc"      , &Image::modify_iptc)
        .def("modify_xmp"       , &Image::modify_xmp)
        .def("modify_raw_xmp"   , &Image::modify_raw_xmp)
        .def("modify_comment"   , &Image::modify_comment)
        .def("modify_icc"       , &Image::modify_icc)
        .def("modify_thumbnail" , &Image::modify_thumbnail)
        .def("clear_exif"       , &Image::clear_exif)
        .def("clear_iptc"       , &Image::clear_iptc)
        .def("clear_xmp"        , &Image::clear_xmp)
        .def("clear_comment"    , &Image::clear_comment)
        .def("clear_icc"        , &Image::clear_icc)
        .def("clear_thumbnail"  , &Image::clear_thumbnail);
}
