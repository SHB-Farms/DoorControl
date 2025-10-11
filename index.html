<script type="text/javascript">
        var gk_isXlsx = false;
        var gk_xlsxFileLookup = {};
        var gk_fileData = {};
        function filledCell(cell) {
          return cell !== '' && cell != null;
        }
        function loadFileData(filename) {
        if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
            try {
                var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON to filter blank rows
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
                // Filter out blank rows (rows where all cells are empty, null, or undefined)
                var filteredData = jsonData.filter(row => row.some(filledCell));

                // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
                var headerRowIndex = filteredData.findIndex((row, index) =>
                  row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
                );
                // Fallback
                if (headerRowIndex === -1 || headerRowIndex > 25) {
                  headerRowIndex = 0;
                }

                // Convert filtered JSON back to CSV
                var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex)); // Create a new sheet from filtered array of arrays
                csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
                return csv;
            } catch (e) {
                console.error(e);
                return "";
            }
        }
        return gk_fileData[filename] || "";
        }
        </script><!DOCTYPE html>
<html>
<head>
    <title>SHB Door Controller</title>
    <style>
        :root {
            --primary: #6366f1;
            --success: #22c55e;
            --danger: #ef4444;
            --warning: #eab308;
            --background: #0f172a;
            --surface: #1e293b;
        }

        body {
            font-family: 'Inter', system-ui, sans-serif;
            margin: 0;
            padding: 40px;
            background: var(--background);
            color: white;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
        }

        .header {
            text-align: center;
            margin-bottom: 4rem;
        }

        .header h1 {
            font-size: 5rem;
        }

        #status {
            font-size: 4rem;
            margin: 3rem 0;
            background: linear-gradient(45deg, #6366f1, #8b5cf6);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-weight: 700;
            text-align: center;
        }

        .button-group {
            display: flex;
            gap: 3rem;
            margin: 3rem 0;
        }

        .control-btn {
            padding: 3rem 6rem;
            font-size: 3rem;
            border: none;
            border-radius: 2rem;
            cursor: pointer;
            transition: all 0.2s ease;
            backdrop-filter: blur(10px);
            background: rgba(255, 255, 255, 0.1);
            color: white;
            border: 2px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 6px 10px -1px rgba(0, 0, 0, 0.1);
        }

        .control-btn:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 20px -3px rgba(0, 0, 0, 0.3);
        }

        #scanBtn {
            background: var(--primary);
            font-size: 2.8rem;
            padding: 2.5rem 5rem;
        }

        #disconnectBtn {
            background: var(--danger);
            font-size: 2.8rem;
            padding: 2.5rem 5rem;
        }

        #openBtn { 
            background: var(--success);
            padding: 3rem 6rem;
        }
        #halfBtn { 
            background: var(--warning);
            padding: 3rem 6rem;
        }
        #closeBtn { 
            background: var(--primary);
            padding: 3rem 6rem;
        }
        #stopBtn { 
            background: var(--danger);
            padding: 8rem 8rem;
            font-size: 8rem;
        }

        .controls {
            display: grid;
            gap: 2.5rem;
            margin-top: 4rem;
            width: 100%;
            max-width: 1200px;
        }

        .connected-door {
            font-size: 3.5rem;
            text-align: center;
            color: #94a3b8;
            margin: 2rem 0;
        }

        .modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.7);
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .modal-content {
            background: var(--surface);
            padding: 2rem;
            border-radius: 1rem;
            text-align: center;
            width: 90%;
            max-width: 400px;
            box-shadow: 0 6px 10px -1px rgba(0, 0, 0, 0.3);
        }

        .modal-content input {
            padding: 1rem;
            font-size: 1.8rem;
            margin: 1rem 0;
            width: 80%;
            border: 2px solid var(--primary);
            border-radius: 0.5rem;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            text-align: center;
        }

        .modal-content button {
            padding: 1rem 2rem;
            font-size: 1.6rem;
            margin: 0.5rem;
            border: none;
            border-radius: 0.5rem;
            cursor: pointer;
            background: var(--primary);
            color: white;
        }

        .modal-content button.cancel {
            background: var(--danger);
        }

        @media (max-width: 768px) {
            .control-btn {
                padding: 2rem 3rem;
                font-size: 2.2rem;
            }
            
            #stopBtn {
                padding: 3rem 6rem;
                font-size: 2.8rem;
            }

            .header h1 {
                font-size: 4rem;
            }

            #status {
                font-size: 3rem;
            }

            .connected-door {
                font-size: 2.5rem;
            }

            .modal-content {
                padding: 1.5rem;
            }

            .modal-content input {
                font-size: 1.5rem;
            }

            .modal-content button {
                font-size: 1.4rem;
            }
        }
    </style>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
</head>
<body>
    <div class="header">
        <h1>SHB Door Control v1.7</h1>
        <div class="connected-door" id="doorStatus">Not connected</div>
    </div>

    <div class="button-group">
        <button class="control-btn" id="scanBtn" onclick="scanDoors()">SCAN DOORS</button>
        <button class="control-btn" id="disconnectBtn" onclick="disconnect()" disabled>DISCONNECT</button>
    </div>

    <div class="controls">
        <button class="control-btn" id="openBtn" onclick="sendCommand('O\n')">OPEN</button>
        <button class="control-btn" id="halfBtn" onclick="sendCommand('H\n')">HALF</button>
        <button class="control-btn" id="closeBtn" onclick="sendCommand('C\n')">CLOSE</button>
        <button class="control-btn" id="stopBtn" onclick="sendCommand('S\n')"> STOP </button>
    </div>

    <div id="status"></div>

    <div class="modal" id="passcodeModal">
        <div class="modal-content">
            <h2 id="modalTitle">Enter Passcode</h2>
            <input type="password" id="passcodeInput" maxlength="4" placeholder="Enter 4-digit code">
            <div>
                <button onclick="verifyPasscode()">Submit</button>
                <button class="cancel" onclick="closePasscodeModal()">Cancel</button>
            </div>
        </div>
    </div>

    <script>
        const UART_SERVICE_UUID = '0000ffe0-0000-1000-8000-00805f9b34fb';
        const UART_CHARACTERISTIC_UUID = '0000ffe1-0000-1000-8000-00805f9b34fb';
        let currentDevice = null;
        let currentCharacteristic = null;
        let connectedDoor = null;
        let selectedDevice = null;

        // ====================== PASSCODE CONFIGURATION ======================
        // EASY TO EDIT: Configure your door passcodes here
        // Use format: 'DoorName': 'Passcode' or 'DoorName': null for no passcode
        const DOOR_PASSCODES = {
            // Examples:
            'D1': '1234',    // Door D1 requires passcode 1234
            'D2': null,      // Door D2 has no passcode (free access)
            'D3': '4045',    // Door D3 requires passcode 4045
            'D4': null,      // Door D4 has no passcode
            'D5': '9999',    // Door D5 requires passcode 9999
            // Add more doors here:
            // 'D6': '0000', // Uncomment and modify as needed
            // 'D7': null,   // Uncomment and modify as needed
        };

        // Default behavior for doors not listed above:
        // Set to a passcode string like '2025' to require passcode by default
        // Set to null to allow free access by default
        const DEFAULT_PASSCODE = null;

        function showPasscodeModal(device) {
            selectedDevice = device;
            const doorName = device.name || 'Unknown';
            document.getElementById('modalTitle').textContent = `Enter Passcode for ${doorName}`;
            document.getElementById('passcodeModal').style.display = 'flex';
            document.getElementById('passcodeInput').value = '';
            document.getElementById('passcodeInput').focus();
        }

        function closePasscodeModal() {
            document.getElementById('passcodeModal').style.display = 'none';
            selectedDevice = null;
            showStatus('Passcode entry cancelled');
        }

        function verifyPasscode() {
            const input = document.getElementById('passcodeInput').value;
            const doorName = selectedDevice?.name || 'Unknown';
            const correctPasscode = DOOR_PASSCODES[doorName] !== undefined ? DOOR_PASSCODES[doorName] : DEFAULT_PASSCODE;

            if (correctPasscode === null) {
                document.getElementById('passcodeModal').style.display = 'none';
                connectToDoor(selectedDevice);
            } else if (input === correctPasscode) {
                document.getElementById('passcodeModal').style.display = 'none';
                connectToDoor(selectedDevice);
            } else {
                showStatus('Incorrect passcode', true);
                document.getElementById('passcodeInput').value = '';
            }
        }

        async function scanDoors() {
            try {
                showStatus('Scanning for doors...');
                const device = await navigator.bluetooth.requestDevice({
                    filters: [{ namePrefix: 'D' }],
                    optionalServices: [UART_SERVICE_UUID]
                });

                // Strict validation for D1-D99
                if (!/^D([1-9]|[1-9][0-9])$/.test(device.name)) {
                    showStatus('Only D1-D99 doors supported', true);
                    return;
                }

                // Check if the door requires a passcode
                const passcode = DOOR_PASSCODES[device.name] !== undefined ? DOOR_PASSCODES[device.name] : DEFAULT_PASSCODE;
                if (passcode === null) {
                    connectToDoor(device);
                } else {
                    showPasscodeModal(device);
                }
            } catch(err) {
                if (err.name !== 'NotFoundError') {
                    showStatus(`Error: ${err.message}`, true);
                } else {
                    showStatus('No Bluetooth devices found', true);
                }
            }
        }

        async function connectToDoor(device) {
            try {
                showStatus(`Connecting to ${device.name}...`);
                if (currentDevice?.gatt.connected) {
                    await currentDevice.gatt.disconnect();
                }

                currentDevice = device;
                const server = await currentDevice.gatt.connect();
                const service = await server.getPrimaryService(UART_SERVICE_UUID);
                currentCharacteristic = await service.getCharacteristic(UART_CHARACTERISTIC_UUID);

                connectedDoor = device.name;
                updateUI();
                showStatus(`Connected to ${connectedDoor}`);

                currentDevice.addEventListener('gattserverdisconnected', () => {
                    connectedDoor = null;
                    updateUI();
                    showStatus('Disconnected', true);
                });

            } catch(err) {
                showStatus(`Connection failed: ${err.message}`, true);
            }
        }

        async function disconnect() {
            if (currentDevice?.gatt.connected) {
                await currentDevice.gatt.disconnect();
            }
            connectedDoor = null;
            updateUI();
            showStatus('Disconnected', true);
        }

        async function sendCommand(command) {
            if (!currentCharacteristic) {
                showStatus('Not connected to any door', true);
                return;
            }

            try {
                await currentCharacteristic.writeValue(new TextEncoder().encode(command));
                showStatus(`Sent: ${command.trim()} command`);
            } catch(err) {
                showStatus(`Send failed: ${err.message}`, true);
            }
        }

        function updateUI() {
            document.getElementById('doorStatus').textContent = 
                connectedDoor ? `Connected to: ${connectedDoor}` : "Not connected";
            document.getElementById('disconnectBtn').disabled = !connectedDoor;
        }

        function showStatus(message, isError = false) {
            const statusDiv = document.getElementById('status');
            statusDiv.innerHTML = message;
            statusDiv.style.color = isError ? '#ef4444' : '#22c55e';
        }
    </script>
</body>
</html>
