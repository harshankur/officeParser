#!/bin/bash

# Array of Node versions to test
VERSIONS=("18" "20" "22" "24")

# Loop through each version
for VERSION in "${VERSIONS[@]}"; do
    echo "=========================================================="
    echo "Testing Node.js version $VERSION"
    echo "=========================================================="
    
    # Build the image for the specific version using the new path for Dockerfile.test
    docker build --build-arg NODE_VERSION=$VERSION -f test/matrix/Dockerfile.test -t officeparser-test:$VERSION .
    
    # Run the container
    docker run --rm officeparser-test:$VERSION
    
    # Check if the test failed
    if [ $? -ne 0 ]; then
        echo "Error: Tests failed for Node.js version $VERSION"
        exit 1
    fi
    
    echo "Successfully tested Node.js version $VERSION"
done

echo "=========================================================="
echo "All tests passed for all versions!"
echo "=========================================================="
