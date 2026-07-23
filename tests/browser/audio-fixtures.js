function createPcm16Wav(options) {
  const config = options || {};
  const sampleRate = Number(config.sampleRate || 48000);
  const channels = Number(config.channels || 1);
  const durationSeconds = Number(config.durationSeconds || 1);
  const frameCount = Math.round(sampleRate * durationSeconds);
  const bytesPerSample = 2;
  const dataBytes = frameCount * channels * bytesPerSample;
  const buffer = Buffer.allocUnsafe(44 + dataBytes);

  buffer.write('RIFF', 0, 'ascii');
  buffer.writeUInt32LE(36 + dataBytes, 4);
  buffer.write('WAVE', 8, 'ascii');
  buffer.write('fmt ', 12, 'ascii');
  buffer.writeUInt32LE(16, 16);
  buffer.writeUInt16LE(1, 20);
  buffer.writeUInt16LE(channels, 22);
  buffer.writeUInt32LE(sampleRate, 24);
  buffer.writeUInt32LE(sampleRate * channels * bytesPerSample, 28);
  buffer.writeUInt16LE(channels * bytesPerSample, 32);
  buffer.writeUInt16LE(16, 34);
  buffer.write('data', 36, 'ascii');
  buffer.writeUInt32LE(dataBytes, 40);

  let offset = 44;
  for (let frame = 0; frame < frameCount; frame += 1) {
    for (let channel = 0; channel < channels; channel += 1) {
      const frequency = 220 + channel * 110;
      const sample = Math.sin(2 * Math.PI * frequency * frame / sampleRate) * 0.22;
      buffer.writeInt16LE(Math.round(sample * 32767), offset);
      offset += bytesPerSample;
    }
  }
  return buffer;
}

function wavFilePayload(options) {
  const config = options || {};
  return {
    name: String(config.name || `fixture-${config.channels || 1}ch.wav`),
    mimeType: 'audio/wav',
    buffer: createPcm16Wav(config)
  };
}

module.exports = { createPcm16Wav, wavFilePayload };
