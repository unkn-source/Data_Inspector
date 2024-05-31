#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/Layout/LayoutCollector.h>
#include <chrono>
#include <filesystem>
#include <FreeImage.h>
#include <iomanip>
#include <iostream>
#include <opencv2.4/opencv2/opencv.hpp>
#include <opencv2/imgcodecs.hpp>
#include <poppler/cpp/poppler-document.h>
#include <sstream>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <sys/stat.h>
#include <ctime>

extern "C" {
    #include <gif_lib.h>
}

using namespace Aspose::Words;

namespace fs = std::filesystem;

const std::string RESET_COLOR = "\033[0m";
const std::string RED_COLOR = "\033[31m";
const std::string GREEN_COLOR = "\033[32m";
const std::string LIGHT_GREEN_COLOR = "\033[92m";

const std::unordered_map<std::string, std::string> FILE_TYPES = {
    {".pdf", "PDF Document"},
    {".doc", "Word Document"},
    {".jpg", "JPEG Image"},
    {".jpeg", "JPEG Image"},
    {".gif", "GIF Image"},
    {".tif", "TIF Image"},
    {".tiff", "TIFF Document"},
};

std::tm stringToTm(const std::string& dateStr) {
    std::tm tm = {};
    std::istringstream ss(dateStr);
    ss >> std::get_time(&tm, "%d/%m/%Y %H:%M:%S");
    if (ss.fail()) {
        throw std::runtime_error("Failed to parse date: " + dateStr);
    }
    return tm;
}

std::string getCreationTime(const std::string& filePath) {
    struct stat attr;
    if (stat(filePath.c_str(), &attr) != 0) {
        throw std::runtime_error("Could not retrieve file attributes.");
    }

    auto mod_time = std::chrono::system_clock::from_time_t(attr.st_mtime);
    std::time_t mod_time_t = std::chrono::system_clock::to_time_t(mod_time);

    std::tm tm_local;
    if (localtime_s(&tm_local, &mod_time_t) != 0) {
        throw std::runtime_error("Could not convert time.");
    }

    std::ostringstream ss;
    ss << std::put_time(&tm_local, "%d/%m/%Y %H:%M:%S");
    return ss.str();
}

bool isWithinRange(const std::tm& date, const std::tm& start, const std::tm& end) {
    auto to_time_t = [](const std::tm& tm) {
        std::tm copy = tm;
        return std::mktime(&copy);
        };

    std::time_t date_time = to_time_t(date);
    std::time_t start_time = to_time_t(start);
    std::time_t end_time = to_time_t(end);

    return (date_time >= start_time && date_time <= end_time);
}

int countDocPages(const std::string& filename) {
    System::SharedPtr<Document> doc = System::MakeObject<Document>(System::String::FromUtf8(filename));

    return doc->get_PageCount();
}

int countTiffImages(const char* filename) {
    FreeImage_Initialise();
    FIMULTIBITMAP* multiBitmap = FreeImage_OpenMultiBitmap(FIF_TIFF, filename, FALSE, TRUE, TRUE, TIFF_DEFAULT);
    if (!multiBitmap) {
        std::cerr << "Could not open TIFF file: " << filename << std::endl;
        FreeImage_DeInitialise();
        return -1;
    }

    int imageCount = FreeImage_GetPageCount(multiBitmap);
    FreeImage_CloseMultiBitmap(multiBitmap, 0);
    FreeImage_DeInitialise();
    return imageCount;
}

double calculateA1SizeWidth(const std::string& imagePath) {
    cv::Mat image = cv::imread(imagePath);

    if (image.empty()) {
        std::cerr << "Error loading the image.\n";
        return 0.0;
    }

    const double dpi = 300.0;
    const double mmPerInch = 25.4;

    double widthPixels = image.cols;
    double heightPixels = image.rows;

    double widthInches = widthPixels / dpi;
    double heightInches = heightPixels / dpi;

    double widthA1 = widthInches * mmPerInch / 841.0;
    double heightA1 = heightInches * mmPerInch / 594.0;

    return static_cast<double>(widthA1);
}

double calculateA1SizeHeight(const std::string& imagePath) {
    cv::Mat image = cv::imread(imagePath);

    const double dpi = 300.0;
    const double mmPerInch = 25.4;

    double widthPixels = image.cols;
    double heightPixels = image.rows;

    double widthInches = widthPixels / dpi;
    double heightInches = heightPixels / dpi;

    double widthA1 = widthInches * mmPerInch / 841.0;
    double heightA1 = heightInches * mmPerInch / 594.0;

    return static_cast<double>(heightA1);
}

double calculateGifA1Width(const std::string& imagePath) {
    GifFileType* gifFile = DGifOpenFileName(imagePath.c_str(), nullptr);
    if (!gifFile) {
        std::cerr << "Error opening the GIF file.\n";
        return -1.0;
    }

    const double dpi = 300.0;
    const double mmPerInch = 25.4;
    const double A1Width_mm = 841.0;

    int widthPixels = gifFile->SWidth;
    double widthInches = static_cast<double>(widthPixels) / dpi;
    double widthA1 = widthInches * mmPerInch / A1Width_mm;

    DGifCloseFile(gifFile, nullptr);

    return widthA1;
}

double calculateGifA1Height(const std::string& imagePath) {
    GifFileType* gifFile = DGifOpenFileName(imagePath.c_str(), nullptr);
    if (!gifFile) {
        std::cerr << "Error opening the GIF file.\n";
        return -1.0;
    }

    const double dpi = 300.0;
    const double mmPerInch = 25.4;
    const double A1Height_mm = 594.0;

    int heightPixels = gifFile->SHeight;
    double heightInches = static_cast<double>(heightPixels) / dpi;
    double heightA1 = heightInches * mmPerInch / A1Height_mm;

    DGifCloseFile(gifFile, nullptr);

    return heightA1;
}

int pageCountPDF(const std::string& pdfFilePath) {
    try {
        auto document = std::unique_ptr<poppler::document>(poppler::document::load_from_file(pdfFilePath));
        
        if (!document || document->is_locked()) {
            std::cerr << "Error opening PDF file or file is password protected.\n";
            return -1;
        }
        return document->pages();
    }
    catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return -1;
    }
}

void scan_directory(const fs::path& path, int level, const std::tm& start, const std::tm& end) {
    for (const auto& entry : fs::directory_iterator(path)) {
        fs::file_status file_stat = entry.status();
        std::string file_type = "Unknown";

        if (fs::is_regular_file(entry.path())) {
            uintmax_t filesize = fs::file_size(entry.path());
            std::string size_unit = "B";

            double size = static_cast<double>(filesize);
            if (size >= 1024 * 1024) {
                size /= (1024 * 1024);
                size_unit = "MB";
            }
            else if (size >= 1024) {
                size /= 1024;
                size_unit = "KB";
            }

            std::string extension = entry.path().extension().string();

            if (!extension.empty() && extension[0] == '.') {
                extension = extension.substr(1);
            }

            auto it = FILE_TYPES.find(extension);
            if (it != FILE_TYPES.end()) {
                file_type = it->second;
            }
            else {
                file_type = extension + " File";
            }

            size_t num_pages = 0;

            double size_in_sheets = 2.5;

            std::string creationTimeStr = getCreationTime(entry.path().string());
            std::tm creationTime = stringToTm(creationTimeStr);

            if (isWithinRange(creationTime, start, end)) {
                std::cout << std::fixed << std::setprecision(2) << std::string(level * 2, ' ')
                    << entry.path().filename().string() << " - " << size << " " << size_unit
                    << " (" << file_type << ")";

                if (extension == "tiff")
                    std::cout << " - " << countTiffImages(entry.path().string().c_str()) << " pages";
                else if (extension == "doc")
                    std::cout << " - " << countDocPages(entry.path().string()) << " pages";
                else if (extension == "pdf")
                    std::cout << " - " << pageCountPDF(entry.path().string()) << " pages";
                else if (extension == "jpg" || extension == "tif")
                    std::cout << " - " << calculateA1SizeWidth(entry.path().string()) << " x " << calculateA1SizeWidth(entry.path().string()) << " A1 sheets";
                else if (extension == "gif")
                    std::cout << " - " << calculateGifA1Width(entry.path().string()) << " x " << calculateGifA1Height(entry.path().string()) << " A1 sheets";
                else
                    std::cout << RED_COLOR << " - " << "Unknown file type" << GREEN_COLOR;
                std::cout << std::endl;
            }
            else {
                std::cout << RED_COLOR << "File: " << entry.path().string() << " creation date is not within the range.\n" << GREEN_COLOR;
            }
        }
    }

    for (const auto& entry : fs::directory_iterator(path)) {
        if (fs::is_directory(entry.path())) {
            std::string directory_name = entry.path().filename().string();
            std::cout << std::string(level * 2, ' ') << directory_name;
            if (level >= 2) {
                std::cout << RED_COLOR << " - ERROR: Level 3 nesting" << GREEN_COLOR << std::endl;
            }
            else {
                std::cout << "/" << std::endl;
                if (level < 2) {
                    scan_directory(entry.path(), level + 1, start, end);
                }
            }
        }
    }
}


int main() {
    std::cout << LIGHT_GREEN_COLOR;

    std::cout << " ______   _______  _______  _______    ___   __    _  _______  _______  _______  _______  _______  _______  ______   \n";
    std::cout << "|      | |   _   ||       ||   _   |  |   | |  |  | ||       ||       ||       ||       ||       ||       ||    _ |  \n";
    std::cout << "|  _    ||  |_|  ||_     _||  |_|  |  |   | |   |_| ||  _____||    _  ||    ___||       ||_     _||   _   ||   | ||  \n";
    std::cout << "| | |   ||       |  |   |  |       |  |   | |       || |_____ |   |_| ||   |___ |       |  |   |  |  | |  ||   |_||_ \n";
    std::cout << "| |_|   ||       |  |   |  |       |  |   | |  _    ||_____  ||    ___||    ___||      _|  |   |  |  |_|  ||    __  |\n";
    std::cout << "|       ||   _   |  |   |  |   _   |  |   | | | |   | _____| ||   |    |   |___ |     |_   |   |  |       ||   |  | |\n";
    std::cout << "|______| |__| |__|  |___|  |__| |__|  |___| |_|  |__||_______||___|    |_______||_______|  |___|  |_______||___|  |_|\n\n";
  
    
    std::cout << GREEN_COLOR;

    std::cout << "--------------------------------------------------------------------------------------------------------------------" << std::endl;
    std::string input_path;
    std::cout << "Enter the path to the directory to scan: ";
    std::getline(std::cin, input_path);

    std::string startStr;
    std::string endStr;

    std::cout << "Enter start date (dd/mm/yyyy HH:MM:SS): ";
    std::getline(std::cin, startStr);

    std::cout << "Enter end date (dd/mm/yyyy HH:MM:SS): ";
    std::getline(std::cin, endStr);

    std::tm start = stringToTm(startStr);
    std::tm end = stringToTm(endStr);

    fs::path path(input_path);
    if (!fs::exists(path) || !fs::is_directory(path)) {
        std::cout << RED_COLOR << "ERROR: The specified path is not a directory or does not exist." << std::endl;
        std::cout << GREEN_COLOR << "--------------------------------------------------------------------------------------------------------------------" << std::endl;
        return 1;
    }

    std::cout << "--------------------------------------------------------------------------------------------------------------------" << std::endl;   
    scan_directory(path, 1, start, end);
    std::cout << "--------------------------------------------------------------------------------------------------------------------" << std::endl;

    std::cout << RESET_COLOR;
    return 0;
}
