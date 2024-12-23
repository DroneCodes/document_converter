# Document Converter

A Java application for converting documents between different formats, including Excel to JSON, Access to JSON, JSON to Excel, and JSON to Access.

## Features

- **Excel to JSON**: Converts Excel files to JSON format
- **Access to JSON**: Converts Access database files to JSON format
- **JSON to Excel**: Converts JSON files to Excel format
- **JSON to Access**: Converts JSON files to Access database format
- **Directory Management**: Automatically handles the creation of necessary directories
- **Error Handling**: Provides detailed error messages for troubleshooting

## Requirements

- Java 21
- Maven
- Dependencies:
    - `com.fasterxml.jackson.core:jackson-databind:2.15.2`
    - `org.apache.poi:poi-ooxml:5.2.3`
    - `net.sf.ucanaccess:ucanaccess:5.0.1`
    - `org.apache.logging.log4j:log4j-core:2.20.0`

## Installation

### Maven

Add this dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>com.fasterxml.jackson.core</groupId>
    <artifactId>jackson-databind</artifactId>
    <version>2.15.2</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.3</version>
</dependency>
<dependency>
    <groupId>net.sf.ucanaccess</groupId>
    <artifactId>ucanaccess</artifactId>
    <version>5.0.1</version>
</dependency>
<dependency>
    <groupId>org.apache.logging.log4j</groupId>
    <artifactId>log4j-core</artifactId>
    <version>2.20.0</version>
</dependency>
```

## Usage

### Running the Application

1. Create an `assets` directory in the project root and add your files there.
2. Run the application:
   ```sh
   mvn exec:java -Dexec.mainClass="org.fisayo.DocumentConverter"
   ```
3. Follow the menu prompts to select the type of conversion and handle the conversion process.

### Example: Convert Excel to JSON

```java
import org.fisayo.DocumentConverter;

public class Main {
    public static void main(String[] args) {
        DocumentConverter.main(new String[]{});
    }
}
```

## Directory Structure

- `assets`: Directory to place input files.
- `results`: Directory where output files will be saved.

## Building from Source

1. Clone the repository:
```bash
git clone https://github.com/DroneCodes/document_converter.git
```

2. Build with Maven:
```bash
cd document-converter
mvn clean install
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

If you encounter any issues or have questions, please file an issue on the GitHub repository.

## Acknowledgments

- Inspired by the need for a versatile document conversion tool
- Built using Java and various libraries for optimal performance
- Designed to handle multiple document formats and conversions