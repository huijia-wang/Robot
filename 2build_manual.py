#!/usr/bin/env python3
"""Safe single-motor sample for DYNAMIXEL XL330-M288-T.

Hardware target:
    PC --USB--> U2D2 --3-pin TTL half-duplex--> XL330-M288-T
    Regulated 5.0 V supply --> U2D2 Power Hub Board --> XL330-M288-T

The program defaults to a small relative move in Position Control Mode and logs
position, velocity, input current, PWM, voltage, temperature and error status.
It never changes EEPROM unless --set-position-mode is explicitly supplied.
"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable


PROTOCOL_VERSION = 2.0
EXPECTED_MODEL_NUMBER = 1200

# XL330-M288-T control table (Protocol 2.0)
ADDR_FIRMWARE_VERSION = 6
ADDR_OPERATING_MODE = 11
ADDR_MIN_POSITION_LIMIT = 52
ADDR_MAX_POSITION_LIMIT = 48
ADDR_TORQUE_ENABLE = 64
ADDR_HARDWARE_ERROR = 70
ADDR_PROFILE_ACCELERATION = 108
ADDR_PROFILE_VELOCITY = 112
ADDR_GOAL_POSITION = 116
ADDR_REALTIME_TICK = 120
ADDR_MOVING = 122
ADDR_MOVING_STATUS = 123
ADDR_PRESENT_PWM = 124
ADDR_PRESENT_CURRENT = 126
ADDR_PRESENT_VELOCITY = 128
ADDR_PRESENT_POSITION = 132
ADDR_PRESENT_INPUT_VOLTAGE = 144
ADDR_PRESENT_TEMPERATURE = 146

POSITION_MODE = 3
PULSES_PER_REV = 4096
POSITION_DEG_PER_PULSE = 360.0 / PULSES_PER_REV
VELOCITY_RPM_PER_UNIT = 0.229
CURRENT_MA_PER_UNIT = 1.0
PWM_PERCENT_PER_UNIT = 0.113
VOLTAGE_V_PER_UNIT = 0.1


def to_signed(value: int, bits: int) -> int:
    """Interpret an unsigned register value as a two's-complement integer."""
    sign_bit = 1 << (bits - 1)
    return value - (1 << bits) if value & sign_bit else value


def decode_hardware_error(value: int) -> str:
    labels = []
    if value & 0x01:
        labels.append("input-voltage")
    if value & 0x04:
        labels.append("overheating")
    if value & 0x10:
        labels.append("electrical-shock/insufficient-power")
    if value & 0x20:
        labels.append("overload")
    unknown = value & ~0x35
    if unknown:
        labels.append(f"unknown-bits-0x{unknown:02X}")
    return "none" if not labels else ";".join(labels)


def utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def run_self_test() -> int:
    assert to_signed(0x0001, 16) == 1
    assert to_signed(0xFFFF, 16) == -1
    assert to_signed(0xFFFFFFFF, 32) == -1
    assert decode_hardware_error(0) == "none"
    assert "overheating" in decode_hardware_error(0x04)
    assert round(2048 * POSITION_DEG_PER_PULSE, 3) == 180.0
    print("Self-test passed: conversions and error decoding are consistent.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Operate and monitor one DYNAMIXEL XL330-M288-T safely."
    )
    parser.add_argument("--port", default="COM3", help="Serial port, e.g. COM3 or /dev/ttyUSB0")
    parser.add_argument("--baud", type=int, default=57600, help="DYNAMIXEL baud rate")
    parser.add_argument("--id", type=int, default=1, help="DYNAMIXEL ID (0-252)")
    parser.add_argument("--duration", type=float, default=5.0, help="Monitor duration in seconds")
    parser.add_argument("--interval", type=float, default=0.1, help="Sampling interval in seconds")
    parser.add_argument(
        "--goal-pulse",
        type=int,
        help="Absolute target in pulses. If omitted, a small relative move is used.",
    )
    parser.add_argument(
        "--delta-deg",
        type=float,
        default=20.0,
        help="Relative move in degrees when --goal-pulse is omitted",
    )
    parser.add_argument(
        "--profile-velocity",
        type=int,
        default=50,
        help="Raw profile velocity; 1 unit = 0.229 rpm",
    )
    parser.add_argument(
        "--profile-acceleration",
        type=int,
        default=10,
        help="Raw profile acceleration; 1 unit = 214.577 rpm^2",
    )
    parser.add_argument(
        "--temp-stop",
        type=float,
        default=65.0,
        help="Software stop threshold in degC (official default limit is 70 degC)",
    )
    parser.add_argument("--min-voltage", type=float, default=3.7, help="Software minimum voltage")
    parser.add_argument("--max-voltage", type=float, default=6.0, help="Software maximum voltage")
    parser.add_argument("--csv", default="xl330_log.csv", help="CSV output path")
    parser.add_argument(
        "--read-only",
        action="store_true",
        help="Only read and log feedback; do not write any register",
    )
    parser.add_argument(
        "--set-position-mode",
        action="store_true",
        help="Allow changing EEPROM Operating Mode(11) to Position Mode(3)",
    )
    parser.add_argument(
        "--allow-model-mismatch",
        action="store_true",
        help="Continue even if ping does not report model number 1200",
    )
    parser.add_argument("--self-test", action="store_true", help="Run offline conversion tests")
    return parser


class DxlDevice:
    def __init__(self, port: Any, packet: Any, dxl_id: int, comm_success: int):
        self.port = port
        self.packet = packet
        self.dxl_id = dxl_id
        self.comm_success = comm_success

    def _check(self, operation: str, comm_result: int, dxl_error: int) -> None:
        if comm_result != self.comm_success:
            detail = self.packet.getTxRxResult(comm_result)
            raise RuntimeError(f"{operation}: communication error: {detail}")
        if dxl_error:
            detail = self.packet.getRxPacketError(dxl_error)
            raise RuntimeError(f"{operation}: DYNAMIXEL error: {detail}")

    def read(self, address: int, size: int) -> int:
        method = {
            1: self.packet.read1ByteTxRx,
            2: self.packet.read2ByteTxRx,
            4: self.packet.read4ByteTxRx,
        }.get(size)
        if method is None:
            raise ValueError(f"Unsupported register size: {size}")
        value, comm_result, dxl_error = method(self.port, self.dxl_id, address)
        self._check(f"read address {address}", comm_result, dxl_error)
        return int(value)

    def write(self, address: int, size: int, value: int) -> None:
        method = {
            1: self.packet.write1ByteTxRx,
            2: self.packet.write2ByteTxRx,
            4: self.packet.write4ByteTxRx,
        }.get(size)
        if method is None:
            raise ValueError(f"Unsupported register size: {size}")
        comm_result, dxl_error = method(self.port, self.dxl_id, address, value)
        self._check(f"write address {address}", comm_result, dxl_error)

    def snapshot(self) -> dict[str, Any]:
        position_raw = to_signed(self.read(ADDR_PRESENT_POSITION, 4), 32)
        velocity_raw = to_signed(self.read(ADDR_PRESENT_VELOCITY, 4), 32)
        current_raw = to_signed(self.read(ADDR_PRESENT_CURRENT, 2), 16)
        pwm_raw = to_signed(self.read(ADDR_PRESENT_PWM, 2), 16)
        voltage_raw = self.read(ADDR_PRESENT_INPUT_VOLTAGE, 2)
        temperature = self.read(ADDR_PRESENT_TEMPERATURE, 1)
        hardware_error = self.read(ADDR_HARDWARE_ERROR, 1)
        return {
            "timestamp_utc": utc_timestamp(),
            "realtime_tick_ms": self.read(ADDR_REALTIME_TICK, 2),
            "position_raw": position_raw,
            "position_deg": position_raw * POSITION_DEG_PER_PULSE,
            "velocity_raw": velocity_raw,
            "velocity_rpm": velocity_raw * VELOCITY_RPM_PER_UNIT,
            "current_raw": current_raw,
            "current_mA": current_raw * CURRENT_MA_PER_UNIT,
            "pwm_raw": pwm_raw,
            "pwm_percent": pwm_raw * PWM_PERCENT_PER_UNIT,
            "input_voltage_V": voltage_raw * VOLTAGE_V_PER_UNIT,
            "temperature_C": temperature,
            "moving": self.read(ADDR_MOVING, 1),
            "moving_status": self.read(ADDR_MOVING_STATUS, 1),
            "hardware_error": hardware_error,
            "hardware_error_text": decode_hardware_error(hardware_error),
        }


CSV_FIELDS = [
    "timestamp_utc",
    "elapsed_s",
    "realtime_tick_ms",
    "position_raw",
    "position_deg",
    "velocity_raw",
    "velocity_rpm",
    "current_raw",
    "current_mA",
    "pwm_raw",
    "pwm_percent",
    "input_voltage_V",
    "temperature_C",
    "moving",
    "moving_status",
    "hardware_error",
    "hardware_error_text",
]


def monitor(
    device: DxlDevice,
    duration: float,
    interval: float,
    csv_path: Path,
    temp_stop: float,
    min_voltage: float,
    max_voltage: float,
    emergency_stop: Callable[[], None] | None,
) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    start = time.monotonic()
    with csv_path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=CSV_FIELDS)
        writer.writeheader()
        while True:
            elapsed = time.monotonic() - start
            if elapsed > duration:
                break
            row = device.snapshot()
            row["elapsed_s"] = round(elapsed, 3)
            writer.writerow(row)
            stream.flush()
            print(
                f"t={elapsed:5.2f}s  pos={row['position_deg']:8.2f}deg  "
                f"vel={row['velocity_rpm']:7.2f}rpm  I={row['current_mA']:7.0f}mA  "
                f"V={row['input_voltage_V']:4.1f}V  T={row['temperature_C']:3d}C  "
                f"err={row['hardware_error_text']}"
            )

            unsafe_reason = None
            if row["temperature_C"] >= temp_stop:
                unsafe_reason = f"temperature {row['temperature_C']}C >= {temp_stop}C"
            elif not (min_voltage <= row["input_voltage_V"] <= max_voltage):
                unsafe_reason = (
                    f"voltage {row['input_voltage_V']:.1f}V outside "
                    f"{min_voltage:.1f}-{max_voltage:.1f}V"
                )
            elif row["hardware_error"]:
                unsafe_reason = f"hardware error: {row['hardware_error_text']}"

            if unsafe_reason:
                if emergency_stop is not None:
                    emergency_stop()
                raise RuntimeError(f"Safety stop: {unsafe_reason}")
            remaining = interval - (time.monotonic() - start - elapsed)
            if remaining > 0:
                time.sleep(remaining)


def validate_args(args: argparse.Namespace) -> None:
    if not 0 <= args.id <= 252:
        raise ValueError("--id must be in 0..252")
    if args.duration <= 0 or args.interval <= 0:
        raise ValueError("--duration and --interval must be positive")
    if not 0 <= args.profile_velocity <= 32767:
        raise ValueError("--profile-velocity must be in 0..32767")
    if not 0 <= args.profile_acceleration <= 32767:
        raise ValueError("--profile-acceleration must be in 0..32767")
    if args.min_voltage >= args.max_voltage:
        raise ValueError("--min-voltage must be lower than --max-voltage")


def main() -> int:
    args = build_parser().parse_args()
    if args.self_test:
        return run_self_test()
    validate_args(args)

    try:
        from dynamixel_sdk import COMM_SUCCESS, PacketHandler, PortHandler
    except ImportError as exc:
        raise SystemExit(
            "dynamixel_sdk is not installed. Run: python -m pip install -r requirements.txt"
        ) from exc

    port_handler = PortHandler(args.port)
    packet_handler = PacketHandler(PROTOCOL_VERSION)
    device = DxlDevice(port_handler, packet_handler, args.id, COMM_SUCCESS)
    port_open = False
    torque_enabled_by_program = False

    def torque_off() -> None:
        nonlocal torque_enabled_by_program
        if not port_open or not torque_enabled_by_program:
            return
        try:
            device.write(ADDR_TORQUE_ENABLE, 1, 0)
            print("Torque disabled.")
        except Exception as exc:  # Best-effort shutdown after a communication fault.
            print(f"WARNING: automatic torque-off failed: {exc}", file=sys.stderr)
            print("Cut the external 5 V motor power manually.", file=sys.stderr)
        finally:
            torque_enabled_by_program = False

    try:
        if not port_handler.openPort():
            raise RuntimeError(f"Cannot open serial port {args.port}")
        port_open = True
        if not port_handler.setBaudRate(args.baud):
            raise RuntimeError(f"Cannot set baud rate {args.baud}")

        model_number, comm_result, dxl_error = packet_handler.ping(port_handler, args.id)
        device._check("ping", comm_result, dxl_error)
        firmware = device.read(ADDR_FIRMWARE_VERSION, 1)
        print(
            f"Connected: port={args.port}, baud={args.baud}, ID={args.id}, "
            f"model={model_number}, firmware={firmware}"
        )
        if model_number != EXPECTED_MODEL_NUMBER and not args.allow_model_mismatch:
            raise RuntimeError(
                f"Expected XL330-M288-T model number {EXPECTED_MODEL_NUMBER}, "
                f"but ping returned {model_number}. Use --allow-model-mismatch only after "
                "checking the other model's control table."
            )

        operating_mode = device.read(ADDR_OPERATING_MODE, 1)
        print(f"Operating Mode(11) = {operating_mode}")

        if args.read_only:
            monitor(
                device,
                args.duration,
                args.interval,
                Path(args.csv),
                args.temp_stop,
                args.min_voltage,
                args.max_voltage,
                emergency_stop=None,
            )
            print(f"Read-only monitoring complete. CSV: {Path(args.csv).resolve()}")
            return 0

        device.write(ADDR_TORQUE_ENABLE, 1, 0)
        if operating_mode != POSITION_MODE:
            if not args.set_position_mode:
                raise RuntimeError(
                    "The motor is not in Position Control Mode(3). Change it with "
                    "DYNAMIXEL Wizard 2.0, or rerun with --set-position-mode to allow "
                    "this EEPROM write."
                )
            device.write(ADDR_OPERATING_MODE, 1, POSITION_MODE)
            operating_mode = device.read(ADDR_OPERATING_MODE, 1)
            if operating_mode != POSITION_MODE:
                raise RuntimeError("Operating Mode verification failed")
            print("Operating Mode changed to Position Control Mode(3).")

        min_position = device.read(ADDR_MIN_POSITION_LIMIT, 4)
        max_position = device.read(ADDR_MAX_POSITION_LIMIT, 4)
        present_position = to_signed(device.read(ADDR_PRESENT_POSITION, 4), 32)
        present_single_turn = present_position % PULSES_PER_REV

        if args.goal_pulse is not None:
            target = args.goal_pulse
        else:
            delta = max(1, round(abs(args.delta_deg) / 360.0 * PULSES_PER_REV))
            direction = 1 if args.delta_deg >= 0 else -1
            candidate = present_single_turn + direction * delta
            if candidate > max_position or candidate < min_position:
                candidate = present_single_turn - direction * delta
            target = min(max(candidate, min_position), max_position)

        if not min_position <= target <= max_position:
            raise ValueError(
                f"Target {target} is outside configured position limits "
                f"{min_position}..{max_position}"
            )

        device.write(ADDR_PROFILE_ACCELERATION, 4, args.profile_acceleration)
        device.write(ADDR_PROFILE_VELOCITY, 4, args.profile_velocity)
        # Preload the current physical angle before torque-on to reduce a startup jump.
        device.write(ADDR_GOAL_POSITION, 4, present_single_turn)
        device.write(ADDR_TORQUE_ENABLE, 1, 1)
        torque_enabled_by_program = True
        time.sleep(0.2)
        device.write(ADDR_GOAL_POSITION, 4, target)
        print(
            f"Move: {present_single_turn} -> {target} pulses "
            f"({present_single_turn * POSITION_DEG_PER_PULSE:.2f} -> "
            f"{target * POSITION_DEG_PER_PULSE:.2f} deg)"
        )

        monitor(
            device,
            args.duration,
            args.interval,
            Path(args.csv),
            args.temp_stop,
            args.min_voltage,
            args.max_voltage,
            emergency_stop=torque_off,
        )
        print(f"Move and monitoring complete. CSV: {Path(args.csv).resolve()}")
        return 0
    except KeyboardInterrupt:
        print("Interrupted by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        torque_off()
        if port_open:
            port_handler.closePort()


if __name__ == "__main__":
    raise SystemExit(main())

